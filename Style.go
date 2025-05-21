package xlwt

import (
	"bytes"
)

// only support default style

var StdNumFormatString = []string{
	"general",
	"0",
	"0.00",
	"#,##0",
	"#,##0.00",
	"\"$\"#,##0_);(\"$\"#,##0)",
	"\"$\"#,##0_);[Red](\"$\"#,##0)",
	"\"$\"#,##0.00_);(\"$\"#,##0.00)",
	"\"$\"#,##0.00_);[Red](\"$\"#,##0.00)",
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
}

const FIRST_USER_DEFINED_NUM_FORMAT_IDX = 164
const DefaultCellXFStyle = 0x11

type XFStyle struct {

	//self.num_format_str  = 'General'
	//self.font            = Formatting.Font()
	//self.alignment       = Formatting.Alignment()
	//self.borders         = Formatting.Borders()
	//self.pattern         = Formatting.Pattern()
	//self.protection      = Formatting.Protection()
}

func (xf *XFStyle) GetBiffData() []byte {
	var buf bytes.Buffer
	// The font with index 4 is omitted in all BIFF versions
	for i := 0; i < 6; i++ {
		buf.Write(DefaultFontRecord())
	}

	buf.Write(DefaultFontRecord())
	buf.Write(NumberFormatRecord(FIRST_USER_DEFINED_NUM_FORMAT_IDX, "General"))

	for i := 0; i < 16; i++ {
		buf.Write(DefaultXFRecord())
	}
	buf.Write(DefaultCellXFRecord())
	buf.Write(CellXFRecord(7))
	buf.Write(StyleRecord())
	return buf.Bytes()
}
