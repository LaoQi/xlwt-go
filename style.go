package xlwt

// Formatting types mirroring the Python xlwt Formatting module. Together with
// the number format string they make up an XFStyle, which is stored in a BIFF8
// XF record.
//
// All of these are plain comparable structs, so they can be used as map keys
// for style deduplication.
//
// Most fields treat their zero value as "unset" and fall back to a documented
// default, so a partially filled style such as
// `XFStyle{Font: Font{Bold: true}}` gives you the default look plus bold text.
//
// Two fields cannot do that, because the file format uses 0 for a meaningful
// non-default setting: Alignment.Vert (0 means top) and Protection.CellUnlocked
// (0 means locked, which is the default anyway). Their zero values are therefore
// used as-is; see the field documentation. Start from NewXFStyle when you want
// the library defaults.

// Font escapement types.
const (
	EscapementNone        = 0x00
	EscapementSuperscript = 0x01
	EscapementSubscript   = 0x02
)

// Font underline types.
const (
	UnderlineNone      = 0x00
	UnderlineSingle    = 0x01
	UnderlineSingleAcc = 0x21
	UnderlineDouble    = 0x02
	UnderlineDoubleAcc = 0x22
)

// Font families.
const (
	FamilyNone       = 0x00
	FamilyRoman      = 0x01
	FamilySwiss      = 0x02
	FamilyModern     = 0x03
	FamilyScript     = 0x04
	FamilyDecorative = 0x05
)

// Font character sets.
const (
	CharsetAnsiLatin       = 0x00
	CharsetSysDefault      = 0x01
	CharsetSymbol          = 0x02
	CharsetAppleRoman      = 0x4D
	CharsetAnsiJapShiftJIS = 0x80
	CharsetAnsiKorHangul   = 0x81
	CharsetAnsiKorJohab    = 0x82
	CharsetAnsiChineseGBK  = 0x86
	CharsetAnsiChineseBig5 = 0x88
	CharsetAnsiGreek       = 0xA1
	CharsetAnsiTurkish     = 0xA2
	CharsetAnsiVietnamese  = 0xA3
	CharsetAnsiHebrew      = 0xB1
	CharsetAnsiArabic      = 0xB2
	CharsetAnsiBaltic      = 0xBA
	CharsetAnsiCyrillic    = 0xCC
	CharsetAnsiThai        = 0xDE
	CharsetAnsiLatinII     = 0xEE
	CharsetOemLatinI       = 0xFF
)

// Colour index meaning "automatic" (the cell foreground/background default).
const ColourAutomatic = 0x7FFF

// Font describes the typeface of a cell.
//
// Zero-value handling: an empty Name becomes "Arial", Height 0 becomes 10pt
// (0x00C8 twips), ColourIndex 0 becomes automatic, Weight 0 becomes 400 and
// Charset 0 becomes the system default.
type Font struct {
	// Height in twips (1/20 point). 0x00C8 is 10 points.
	Height int
	// Italic renders the text in italics.
	Italic bool
	// StrikeOut draws a line through the text.
	StrikeOut bool
	// Outline draws the text outlined.
	Outline bool
	// Shadow draws the text with a shadow.
	Shadow bool
	// ColourIndex is a palette index, or ColourAutomatic.
	ColourIndex int
	// Bold renders the text bold. Bold also forces the record weight to 700.
	Bold bool
	// Weight is the font weight (100-1000); 400 is normal, 700 is bold.
	Weight int
	// Escapement is one of the Escapement* constants.
	Escapement int
	// Underline is one of the Underline* constants.
	Underline int
	// Family is one of the Family* constants.
	Family int
	// Charset is one of the Charset* constants.
	Charset int
	// Name is the font name, for example "Arial".
	Name string
}

// NewFont returns the default font: Arial 10 point, automatic colour, weight
// 400, no bold/italic/underline.
func NewFont() Font {
	return defaultFont()
}

func defaultFont() Font {
	return Font{
		Height:      0x00C8,
		ColourIndex: ColourAutomatic,
		Weight:      0x0190,
		Escapement:  EscapementNone,
		Underline:   UnderlineNone,
		Family:      FamilyNone,
		Charset:     CharsetSysDefault,
		Name:        "Arial",
	}
}

// withDefaults returns the font with zero values replaced by their defaults.
func (f Font) withDefaults() Font {
	if f.Height == 0 {
		f.Height = 0x00C8
	}
	if f.ColourIndex == 0 {
		f.ColourIndex = ColourAutomatic
	}
	if f.Weight == 0 {
		f.Weight = 0x0190
	}
	if f.Charset == 0 {
		f.Charset = CharsetSysDefault
	}
	if f.Name == "" {
		f.Name = "Arial"
	}
	return f
}

// effectiveWeight is the weight actually written: bold forces 700, exactly like
// upstream xlwt.
func (f Font) effectiveWeight() int {
	if f.Bold {
		return 0x02BC
	}
	return f.Weight
}

// options packs the font option flags.
func (f Font) options() int {
	options := 0x00
	if f.Bold {
		options |= 0x01
	}
	if f.Italic {
		options |= 0x02
	}
	if f.Underline != UnderlineNone {
		options |= 0x04
	}
	if f.StrikeOut {
		options |= 0x08
	}
	if f.Outline {
		options |= 0x10
	}
	if f.Shadow {
		options |= 0x20
	}
	return options
}

// Horizontal alignment values.
const (
	HorzGeneral            = 0x00
	HorzLeft               = 0x01
	HorzCenter             = 0x02
	HorzRight              = 0x03
	HorzFilled             = 0x04
	HorzJustified          = 0x05
	HorzCenterAcrossSelect = 0x06
	HorzDistributed        = 0x07
)

// Vertical alignment values.
const (
	VertTop         = 0x00
	VertCenter      = 0x01
	VertBottom      = 0x02
	VertJustified   = 0x03
	VertDistributed = 0x04
)

// Text direction values.
const (
	DirectionGeneral = 0x00
	DirectionLR      = 0x01
	DirectionRL      = 0x02
)

// Text rotation values; any angle between 1 and 180 is also valid.
const (
	RotationNone    = 0x00
	RotationStacked = 0xFF
)

// Alignment describes how the cell content is positioned.
//
// The zero value aligns to the top, because 0 is the file format code for
// "top". The library default is bottom aligned, so start from NewAlignment (or
// from NewXFStyle) when you want the usual look.
type Alignment struct {
	// Horz is one of the Horz* constants.
	Horz int
	// Vert is one of the Vert* constants.
	Vert int
	// Dire is one of the Direction* constants.
	Dire int
	// Rota is a rotation angle: 0 none, 1..90 counter-clockwise,
	// 91..180 clockwise, RotationStacked for stacked text.
	Rota int
	// Wrap wraps the text at the right border.
	Wrap bool
	// Shrink shrinks the content to fit into the cell.
	Shrink bool
	// Indent is the indentation level (0..15).
	Indent int
	// Merge is the "merge" attribute of the XF record.
	Merge bool
}

// NewAlignment returns the default alignment: general horizontal, bottom
// vertical, no wrapping, no indentation.
func NewAlignment() Alignment {
	return defaultAlignment()
}

func defaultAlignment() Alignment {
	return Alignment{Vert: VertBottom}
}

// withDefaults returns the alignment unchanged: every field of Alignment has a
// meaningful zero value.
func (a Alignment) withDefaults() Alignment {
	return a
}

// wrapFlag packs the wrap attribute as a bit.
func (a Alignment) wrapFlag() int {
	if a.Wrap {
		return 1
	}
	return 0
}

func (a Alignment) shrinkFlag() int {
	if a.Shrink {
		return 1
	}
	return 0
}

func (a Alignment) mergeFlag() int {
	if a.Merge {
		return 1
	}
	return 0
}

// Border line styles.
const (
	BorderNone                    = 0x00
	BorderThin                    = 0x01
	BorderMedium                  = 0x02
	BorderDashed                  = 0x03
	BorderDotted                  = 0x04
	BorderThick                   = 0x05
	BorderDouble                  = 0x06
	BorderHair                    = 0x07
	BorderMediumDashed            = 0x08
	BorderThinDashDotted          = 0x09
	BorderMediumDashDotted        = 0x0A
	BorderThinDashDotDotted       = 0x0B
	BorderMediumDashDotDotted     = 0x0C
	BorderSlantedMediumDashDotted = 0x0D
)

// Border colour default.
const ColourBorderAutomatic = 0x40

// Borders describes the cell border lines.
//
// Zero-value handling: colours left at 0 become the automatic border colour;
// colours of a BorderNone line are always written as 0.
type Borders struct {
	Left, Right, Top, Bottom, Diag int
	// Colours, one per line, or 0 for the automatic border colour.
	LeftColour, RightColour, TopColour, BottomColour, DiagColour int
	// NeedDiag1/NeedDiag2 draw the diagonals from top-left to bottom-right and
	// from bottom-left to top-right respectively.
	NeedDiag1 bool
	NeedDiag2 bool
}

// NewBorders returns borders without any line.
func NewBorders() Borders {
	return defaultBorders()
}

func defaultBorders() Borders {
	return Borders{
		Left:         BorderNone,
		Right:        BorderNone,
		Top:          BorderNone,
		Bottom:       BorderNone,
		Diag:         BorderNone,
		LeftColour:   ColourBorderAutomatic,
		RightColour:  ColourBorderAutomatic,
		TopColour:    ColourBorderAutomatic,
		BottomColour: ColourBorderAutomatic,
		DiagColour:   ColourBorderAutomatic,
	}
}

// withDefaults returns the borders with zero colours replaced by the automatic
// border colour.
func (b Borders) withDefaults() Borders {
	if b.LeftColour == 0 {
		b.LeftColour = ColourBorderAutomatic
	}
	if b.RightColour == 0 {
		b.RightColour = ColourBorderAutomatic
	}
	if b.TopColour == 0 {
		b.TopColour = ColourBorderAutomatic
	}
	if b.BottomColour == 0 {
		b.BottomColour = ColourBorderAutomatic
	}
	if b.DiagColour == 0 {
		b.DiagColour = ColourBorderAutomatic
	}
	return b
}

// Pattern values. 0x00 is no fill, 0x01 a solid fill; 0x02..0x12 are the
// remaining BIFF8 fill patterns.
const (
	PatternNone  = 0x00
	PatternSolid = 0x01
)

// Pattern colour defaults.
const (
	ColourPatternForeground = 0x40
	ColourPatternBackground = 0x41
)

// Pattern describes the background fill of a cell.
//
// Zero-value handling: foreground 0 becomes the automatic fill colour and
// background 0 becomes the automatic background colour.
type Pattern struct {
	// Pattern is one of the Pattern* constants.
	Pattern int
	// ForeColour is the fill (pattern) colour index.
	ForeColour int
	// BackColour is the background colour index.
	BackColour int
}

// NewPattern returns a pattern without a fill.
func NewPattern() Pattern {
	return defaultPattern()
}

func defaultPattern() Pattern {
	return Pattern{
		Pattern:    PatternNone,
		ForeColour: ColourPatternForeground,
		BackColour: ColourPatternBackground,
	}
}

// withDefaults returns the pattern with zero colours replaced by defaults.
func (p Pattern) withDefaults() Pattern {
	if p.ForeColour == 0 {
		p.ForeColour = ColourPatternForeground
	}
	if p.BackColour == 0 {
		p.BackColour = ColourPatternBackground
	}
	return p
}

// Protection describes cell locking.
//
// The zero value means "locked", which is the file format default, so an
// unlocked cell has to be asked for explicitly. Locking only has an effect once
// the sheet itself is protected.
type Protection struct {
	// CellUnlocked unlocks the cell. The zero value keeps it locked.
	CellUnlocked bool
	// FormulaHidden hides the formula of the cell.
	FormulaHidden bool
}

// NewProtection returns the default protection: a locked cell.
func NewProtection() Protection {
	return Protection{}
}

func defaultProtection() Protection {
	return Protection{}
}

func (p Protection) cellLockedFlag() int {
	if p.CellUnlocked {
		return 0
	}
	return 1
}

func (p Protection) formulaHiddenFlag() int {
	if p.FormulaHidden {
		return 1
	}
	return 0
}

// XFStyle is the collection of formatting attributes of a cell: a number format
// string plus the five formatting components.
//
// Most zero values fall back to the defaults described on the individual
// components, so a partially filled style works as expected. The exceptions are
// Alignment.Vert and Protection.CellUnlocked, whose zero values are meaningful
// settings (see their documentation). Use NewXFStyle to start from an explicit
// set of defaults.
type XFStyle struct {
	// NumFormatStr is an Excel number format string such as "General", "0.00"
	// or "#,##0". Empty means "General".
	NumFormatStr string
	Font         Font
	Alignment    Alignment
	Borders      Borders
	Pattern      Pattern
	Protection   Protection
}

// NewXFStyle returns a style equal to the library default: Arial 10pt, general
// number format, bottom-aligned, no borders, no fill, locked cell.
func NewXFStyle() *XFStyle {
	align := defaultAlignment()
	prot := defaultProtection()
	return &XFStyle{
		NumFormatStr: "General",
		Font:         defaultFont(),
		Alignment:    align,
		Borders:      defaultBorders(),
		Pattern:      defaultPattern(),
		Protection:   prot,
	}
}

// withDefaults fills in zero values, so that partially initialised styles work.
func (s XFStyle) withDefaults() XFStyle {
	if s.NumFormatStr == "" {
		s.NumFormatStr = "General"
	}
	s.Font = s.Font.withDefaults()
	s.Alignment = s.Alignment.withDefaults()
	s.Borders = s.Borders.withDefaults()
	s.Pattern = s.Pattern.withDefaults()
	return s
}

// stdNumFormatStrings are the built-in number formats. Their position in this
// list determines their BIFF8 format index (see standardNumFormats), so the
// order must not change.
var stdNumFormatStrings = []string{
	"general",
	"0",
	"0.00",
	"#,##0",
	"#,##0.00",
	`"$"#,##0_);("$"#,##0)`,
	`"$"#,##0_);[Red]("$"#,##0)`,
	`"$"#,##0.00_);("$"#,##0.00)`,
	`"$"#,##0.00_);[Red]("$"#,##0.00)`,
	"0%",
	"0.00%",
	"0.00E+00",
	"# ?/?",
	"# ??/??",
	"M/D/YY",
	"D-MMM-YY",
	"D-MMM",
	"MMM-YY",
	"h:mm AM/PM",
	"h:mm:ss AM/PM",
	"h:mm",
	"h:mm:ss",
	"M/D/YY h:mm",
	"_(#,##0_);(#,##0)",
	"_(#,##0_);[Red](#,##0)",
	"_(#,##0.00_);(#,##0.00)",
	"_(#,##0.00_);[Red](#,##0.00)",
	`_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)`,
	`_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)`,
	`_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)`,
	`_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)`,
	"mm:ss",
	"[h]:mm:ss",
	"mm:ss.0",
	"##0.0E+0",
	"@",
}

// FIRST_USER_DEFINED_NUM_FORMAT_IDX is the first format index not reserved by
// the file format; custom format strings are numbered from here upwards.
const FIRST_USER_DEFINED_NUM_FORMAT_IDX = 164

// DefaultCellXFStyle is the XF index used for cells written without an explicit
// style. It matches the index the Python original assigns to its default style.
const DefaultCellXFStyle = 0x11

// standardNumFormats builds the format string to index mapping of the built-in
// number formats: the first 23 entries are indices 0..22, the remaining 13 are
// indices 37..49 (24..36 are reserved by the file format).
func standardNumFormats() map[string]int {
	formats := make(map[string]int, len(stdNumFormatStrings))
	for i, s := range stdNumFormatStrings[:23] {
		formats[s] = i
	}
	for i, s := range stdNumFormatStrings[23:] {
		formats[s] = 37 + i
	}
	return formats
}
