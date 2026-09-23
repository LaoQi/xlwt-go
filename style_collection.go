package xlwt

import (
	"bytes"
	"fmt"
	"sort"
)

// maxCellXFIndex is the highest XF index a cell may reference. 0xFFF is reserved
// as a sentinel value, so 4094 cell XFs (0x10..0xFFE) are available.
const maxCellXFIndex = 0xFFE

// styleComponent is the part of a style an XF record points at besides the font
// and the number format.
type styleComponent struct {
	Alignment  Alignment
	Borders    Borders
	Pattern    Pattern
	Protection Protection
}

// defaultStyleComponent is the component of the library default style.
func defaultStyleComponent() styleComponent {
	return styleComponent{
		Alignment:  defaultAlignment(),
		Borders:    defaultBorders(),
		Pattern:    defaultPattern(),
		Protection: defaultProtection(),
	}
}

// fontEntry identifies a font by value plus its font index.
type fontEntry struct {
	index int
	font  Font
}

// xfEntry is the full description of an XF record.
type xfEntry struct {
	fontIndex int
	numFormat int
	component styleComponent
}

// styleCollection assigns font, number format and XF indexes to styles and
// serialises the style section of the workbook globals stream.
//
// Index allocation reproduces the Python original so that the default workbook
// produces byte-identical output:
//
//   - fonts 0, 1, 2, 3 and 5 hold the default font (index 4 is skipped in every
//     BIFF version), font 6 belongs to the library's internal default style and
//     font 7 is the one user styles share;
//   - sixteen style XFs occupy indexes 0x00..0x0F and all point at font 6;
//   - cell XFs start at 0x10; the library default is therefore 0x11, which is
//     what DefaultCellXFStyle documents.
//
// Beyond that, styles and fonts are deduplicated by value: writing the same
// style to a million cells registers it once. This is stricter than the Python
// original, which deduplicates by object identity and therefore records a new
// XF for every equivalent style object.
type styleCollection struct {
	fonts         map[int]Font
	fontIndex     map[Font]int
	nextFontIndex int

	numFormats map[string]int

	xfEntries []xfEntry
	xfIndex   map[xfEntry]int
}

// NewStyleCollection is the constructor used internally; it is not exported.
func newStyleCollection() *styleCollection {
	sc := &styleCollection{
		fonts:         make(map[int]Font),
		fontIndex:     make(map[Font]int),
		nextFontIndex: 8,
		numFormats:    standardNumFormats(),
		xfIndex:       make(map[xfEntry]int),
	}

	// The font with index 4 is omitted in all BIFF versions.
	for _, idx := range []int{0, 1, 2, 3, 5, 6, 7} {
		sc.fonts[idx] = defaultFont()
	}
	// User styles share the font of the public default style; font 6 is only
	// referenced by the built-in style XFs.
	sc.fontIndex[defaultFont()] = 7

	// The internal default style becomes cell XF 0x10, the public default style
	// becomes 0x11. Upstream xlwt ends up with both because they are separate
	// objects, and the file layout depends on it.
	sc.numFormats["General"] = FIRST_USER_DEFINED_NUM_FORMAT_IDX
	sc.appendXF(xfEntry{fontIndex: 6, numFormat: FIRST_USER_DEFINED_NUM_FORMAT_IDX, component: defaultStyleComponent()})
	sc.appendXF(xfEntry{fontIndex: 7, numFormat: FIRST_USER_DEFINED_NUM_FORMAT_IDX, component: defaultStyleComponent()})

	return sc
}

// appendXF registers an XF entry and returns its index.
func (sc *styleCollection) appendXF(entry xfEntry) int {
	idx := 0x10 + len(sc.xfEntries)
	sc.xfEntries = append(sc.xfEntries, entry)
	sc.xfIndex[entry] = idx
	return idx
}

// AddStyle registers a style and returns the XF index a cell should use. A nil
// style means the library default. Registering the same style twice returns the
// same index.
func (sc *styleCollection) AddStyle(style *XFStyle) (int, error) {
	resolved := NewXFStyle()
	if style != nil {
		s := style.withDefaults()
		resolved = &s
	}

	numFormat, err := sc.addNumFormat(resolved.NumFormatStr)
	if err != nil {
		return 0, err
	}
	fontIdx := sc.addFont(resolved.Font)

	entry := xfEntry{
		fontIndex: fontIdx,
		numFormat: numFormat,
		component: styleComponent{
			Alignment:  resolved.Alignment,
			Borders:    resolved.Borders,
			Pattern:    resolved.Pattern,
			Protection: resolved.Protection,
		},
	}
	if idx, ok := sc.xfIndex[entry]; ok {
		return idx, nil
	}

	// fsanity: respect the limit before touching the state
	if 0x10+len(sc.xfEntries) > maxCellXFIndex {
		return 0, fmt.Errorf("%w: at most %d distinct cell styles", ErrTooManyStyles, maxCellXFIndex-0x10+1)
	}
	return sc.appendXF(entry), nil
}

// addFont returns the index of font, registering it if necessary.
func (sc *styleCollection) addFont(font Font) int {
	if idx, ok := sc.fontIndex[font]; ok {
		return idx
	}
	idx := sc.nextFontIndex
	// Index 4 is never used.
	if idx == 4 {
		idx = 5
	}
	for {
		if _, taken := sc.fonts[idx]; !taken {
			break
		}
		idx++
		if idx == 4 {
			idx = 5
		}
	}
	sc.nextFontIndex = idx + 1
	sc.fontIndex[font] = idx
	sc.fonts[idx] = font
	return idx
}

// addNumFormat returns the format index of format, registering a user defined
// format if it is not one of the built-in ones.
func (sc *styleCollection) addNumFormat(format string) (int, error) {
	if idx, ok := sc.numFormats[format]; ok {
		return idx, nil
	}
	idx := FIRST_USER_DEFINED_NUM_FORMAT_IDX + sc.customNumFormatCount()
	if idx > 0xFFFF {
		return 0, fmt.Errorf("%w: %q", ErrTooManyNumberFormats, format)
	}
	sc.numFormats[format] = idx
	return idx, nil
}

// customNumFormatCount counts the registered user defined number formats.
func (sc *styleCollection) customNumFormatCount() int {
	n := 0
	for _, idx := range sc.numFormats {
		if idx >= FIRST_USER_DEFINED_NUM_FORMAT_IDX {
			n++
		}
	}
	return n
}

// GetBiffData serialises the style section: fonts, user defined number formats,
// the sixteen built-in style XFs, the cell XFs and the STYLE record.
func (sc *styleCollection) GetBiffData() []byte {
	var buf bytes.Buffer

	for _, idx := range sortedFontIndexes(sc.fonts) {
		buf.Write(fontRecord(sc.fonts[idx]))
	}

	// Only user defined formats need a FORMAT record; the built-in ones are
	// implicit.
	for _, idx := range sc.customNumFormatIndexes() {
		buf.Write(numberFormatRecord(idx, sc.numFormatString(idx)))
	}

	// Sixteen built-in style XFs, all pointing at font 6.
	styleEntry := xfEntry{
		fontIndex: 6,
		numFormat: FIRST_USER_DEFINED_NUM_FORMAT_IDX,
		component: defaultStyleComponent(),
	}
	for i := 0; i < 16; i++ {
		buf.Write(styleEntry.record(true))
	}

	for _, entry := range sc.xfEntries {
		buf.Write(entry.record(false))
	}

	buf.Write(styleRecord())
	return buf.Bytes()
}

// record serialises the XF entry.
func (e xfEntry) record(styleXF bool) []byte {
	return xfRecord(e.fontIndex, e.numFormat, e.component.Alignment,
		e.component.Borders, e.component.Pattern, e.component.Protection, styleXF)
}

// numFormatString returns the format string registered under idx.
func (sc *styleCollection) numFormatString(idx int) string {
	for s, i := range sc.numFormats {
		if i == idx {
			return s
		}
	}
	return "General"
}

// customNumFormatIndexes returns the user defined format indexes in order.
func (sc *styleCollection) customNumFormatIndexes() []int {
	var idxs []int
	for _, idx := range sc.numFormats {
		if idx >= FIRST_USER_DEFINED_NUM_FORMAT_IDX {
			idxs = append(idxs, idx)
		}
	}
	sort.Ints(idxs)
	return idxs
}

// sortedFontIndexes returns font indexes in order.
func sortedFontIndexes(fonts map[int]Font) []int {
	idxs := make([]int, 0, len(fonts))
	for idx := range fonts {
		idxs = append(idxs, idx)
	}
	sort.Ints(idxs)
	return idxs
}
