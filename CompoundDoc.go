package xlwt

import (
	"bytes"
	"encoding/binary"
	"io"
	"log"
	"unicode/utf16"
)

type XlsDoc struct {
	buf              bytes.Buffer
	book_stream_len  int
	book_stream_sect []int
	dir_stream       []byte
	dir_stream_sect  []int
	packed_SAT       []byte
	SAT_sect         []int
	packed_MSAT_1st  []byte
	packed_MSAT_2nd  []byte
	MSAT_sect_2nd    []int
}

func NewXlsDoc() *XlsDoc {
	return &XlsDoc{
		buf:              bytes.Buffer{},
		book_stream_len:  0,
		book_stream_sect: make([]int, 0),
		dir_stream:       make([]byte, SECTOR_SIZE),
		dir_stream_sect:  make([]int, 0),
		packed_SAT:       make([]byte, 0),
		SAT_sect:         make([]int, 0),
		packed_MSAT_1st:  make([]byte, 0),
		packed_MSAT_2nd:  make([]byte, 0),
		MSAT_sect_2nd:    make([]int, 0),
	}
}

func (self *XlsDoc) BuildDirectory() {
	var buf bytes.Buffer
	name := utf16.Encode([]rune("Root Entry\x00"))
	var nameBuf bytes.Buffer
	_ = binary.Write(&nameBuf, binary.LittleEndian, name)
	nameBytes := nameBuf.Bytes()

	namePaddingLength := 64 - len(nameBytes)
	namePadding := make([]byte, namePaddingLength)
	_ = binary.Write(&buf, binary.LittleEndian, nameBytes)
	_ = binary.Write(&buf, binary.LittleEndian, namePadding)
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(len(nameBytes)))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x05)) // dentry_type = 0x05
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x01)) // dentry_colour    = 0x01
	_ = binary.Write(&buf, binary.LittleEndian, SP_l(-1))   // dentry_did_left = -1
	_ = binary.Write(&buf, binary.LittleEndian, SP_l(-1))   // dentry_did_right = -1
	_ = binary.Write(&buf, binary.LittleEndian, SP_l(1))    // dentry_did_root = 1

	for i := 0; i < 9; i++ {
		_ = binary.Write(&buf, binary.LittleEndian, SP_L(0))
	}
	_ = binary.Write(&buf, binary.LittleEndian, SP_l(-2)) // dentry_start_sid = 1
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(0))  // dentry_stream_sz = 0
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(0))

	// Workbook
	name = utf16.Encode([]rune("Workbook\x00"))
	nameBuf = bytes.Buffer{}
	_ = binary.Write(&nameBuf, binary.LittleEndian, name)
	nameBytes = nameBuf.Bytes()

	namePaddingLength = 64 - len(nameBytes)
	namePadding = make([]byte, namePaddingLength)
	_ = binary.Write(&buf, binary.LittleEndian, nameBytes)
	_ = binary.Write(&buf, binary.LittleEndian, namePadding)
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(len(nameBytes)))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x02)) // dentry_type = 0x02
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x01)) // dentry_colour    = 0x01
	_ = binary.Write(&buf, binary.LittleEndian, SP_l(-1))   // dentry_did_left = -1
	_ = binary.Write(&buf, binary.LittleEndian, SP_l(-1))   // dentry_did_right = -1
	_ = binary.Write(&buf, binary.LittleEndian, SP_l(-1))   // dentry_did_root = -1

	for i := 0; i < 9; i++ {
		_ = binary.Write(&buf, binary.LittleEndian, SP_L(0))
	}
	_ = binary.Write(&buf, binary.LittleEndian, SP_l(0))                    // dentry_start_sid = 0
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(self.book_stream_len)) // dentry_stream_sz = 0
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(0))

	// padding
	padding := make([]byte, 64)
	_ = binary.Write(&buf, binary.LittleEndian, padding)
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x00)) // dentry_type = 0x00 # empty
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x01)) // dentry_colour    = 0x01
	_ = binary.Write(&buf, binary.LittleEndian, SP_l(-1))   // dentry_did_left = -1
	_ = binary.Write(&buf, binary.LittleEndian, SP_l(-1))   // dentry_did_right = -1
	_ = binary.Write(&buf, binary.LittleEndian, SP_l(-1))   // dentry_did_root = -1

	for i := 0; i < 9; i++ {
		_ = binary.Write(&buf, binary.LittleEndian, SP_L(0))
	}
	_ = binary.Write(&buf, binary.LittleEndian, SP_l(-2)) // dentry_start_sid = -2
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(0))  // dentry_stream_sz = 0
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(0))

	// padding
	padding = make([]byte, 64)
	_ = binary.Write(&buf, binary.LittleEndian, padding)
	_ = binary.Write(&buf, binary.LittleEndian, SP_H(0))
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x00)) // dentry_type = 0x00 # empty
	_ = binary.Write(&buf, binary.LittleEndian, SP_B(0x01)) // dentry_colour    = 0x01
	_ = binary.Write(&buf, binary.LittleEndian, SP_l(-1))   // dentry_did_left = -1
	_ = binary.Write(&buf, binary.LittleEndian, SP_l(-1))   // dentry_did_right = -1
	_ = binary.Write(&buf, binary.LittleEndian, SP_l(-1))   // dentry_did_root = -1

	for i := 0; i < 9; i++ {
		_ = binary.Write(&buf, binary.LittleEndian, SP_L(0))
	}
	_ = binary.Write(&buf, binary.LittleEndian, SP_l(-2)) // dentry_start_sid = -2
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(0))  // dentry_stream_sz = 0
	_ = binary.Write(&buf, binary.LittleEndian, SP_L(0))

	self.dir_stream = buf.Bytes()
}

func (self *XlsDoc) BuildSat() {
	book_sect_count := self.book_stream_len >> 9
	dir_sect_count := len(self.dir_stream) >> 9

	total_sect_count := book_sect_count + dir_sect_count
	SAT_sect_count := 0
	MSAT_sect_count := 0
	SAT_sect_count_limit := 109

	for total_sect_count > 128*SAT_sect_count || SAT_sect_count > SAT_sect_count_limit {
		SAT_sect_count += 1
		total_sect_count += 1
		if SAT_sect_count > SAT_sect_count_limit {
			MSAT_sect_count += 1
			total_sect_count += 1
			SAT_sect_count_limit += 127
		}
	}

	SAT := FillInt(128*SAT_sect_count, SID_FREE_SECTOR)

	sect := 0
	for sect < book_sect_count-1 {
		self.book_stream_sect = append(self.dir_stream_sect, sect)
		SAT[sect] = sect + 1
		sect += 1
	}

	self.book_stream_sect = append(self.book_stream_sect, sect)
	SAT[sect] = SID_END_OF_CHAIN
	sect += 1

	for sect < book_sect_count+MSAT_sect_count {
		self.MSAT_sect_2nd = append(self.MSAT_sect_2nd, sect)
		SAT[sect] = SID_USED_BY_MSAT
		sect += 1
	}

	for sect < book_sect_count+MSAT_sect_count+SAT_sect_count {
		self.SAT_sect = append(self.SAT_sect, sect)
		SAT[sect] = SID_USED_BY_SAT
		sect += 1
	}

	for sect < book_sect_count+MSAT_sect_count+SAT_sect_count+dir_sect_count-1 {
		self.dir_stream_sect = append(self.dir_stream_sect, sect)
		SAT[sect] = sect + 1
		sect += 1
	}
	self.dir_stream_sect = append(self.dir_stream_sect, sect)
	SAT[sect] = SID_END_OF_CHAIN
	sect += 1

	var packedSATBuffer bytes.Buffer
	for i := 0; i < len(SAT); i++ {
		_ = binary.Write(&packedSATBuffer, binary.LittleEndian, SP_l(SAT[i]))
	}
	self.packed_SAT = packedSATBuffer.Bytes()

	MSAT_1st := FillInt(109, SID_FREE_SECTOR)
	for i := 0; i < 109 && i < len(self.SAT_sect); i++ {
		MSAT_1st[i] = self.SAT_sect[i]
	}

	var packedMSAT1stBuffer bytes.Buffer
	for i := 0; i < len(MSAT_1st); i++ {
		_ = binary.Write(&packedMSAT1stBuffer, binary.LittleEndian, SP_l(MSAT_1st[i]))
	}
	self.packed_MSAT_1st = packedMSAT1stBuffer.Bytes()

	MSAT_2nd := FillInt(128*MSAT_sect_count, SID_FREE_SECTOR)
	if MSAT_sect_count > 0 {
		MSAT_2nd[len(MSAT_2nd)-1] = SID_END_OF_CHAIN
	}

	msat_sect := 0
	sid_num := 0
	for i := 109; i < SAT_sect_count; {
		if (sid_num+1)%128 == 0 {
			msat_sect += 1
			if msat_sect < len(self.MSAT_sect_2nd) {
				MSAT_2nd[sid_num] = self.MSAT_sect_2nd[msat_sect]
			}
		} else {
			MSAT_2nd[sid_num] = self.SAT_sect[i]
		}
		sid_num += 1
	}
	var packedMSAT2ndBuffer bytes.Buffer
	for i := 0; i < len(MSAT_2nd); i++ {
		_ = binary.Write(&packedMSAT2ndBuffer, binary.LittleEndian, SP_l(MSAT_2nd[i]))
	}
	self.packed_MSAT_2nd = packedMSAT2ndBuffer.Bytes()
}

func (self *XlsDoc) WriteHeader() {
	// 主文件头
	header := []byte{
		0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, // 文件头魔数
		0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // 保留字段
		0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // 保留字段
		0x3E, 0x00, // rev_num
		0x03, 0x00, // ver_num
		0xFE, 0xFF, // byte_order
		0x09, 0x00, // log_sect_size
		0x06, 0x00, // log_short_sect_size
		0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // 保留字段
		0x00, 0x00, // 保留字段
	}

	self.buf.Write(header)
	_ = binary.Write(&self.buf, binary.LittleEndian, SP_L(len(self.SAT_sect)))
	_ = binary.Write(&self.buf, binary.LittleEndian, SP_l(self.dir_stream_sect[0])) //  dir_start_sid = struct.pack('<l', self.dir_stream_sect[0])
	self.buf.Write([]byte{0x00, 0x00, 0x00, 0x00})
	_ = binary.Write(&self.buf, binary.LittleEndian, SP_L(0x1000))
	_ = binary.Write(&self.buf, binary.LittleEndian, SP_l(-2))
	_ = binary.Write(&self.buf, binary.LittleEndian, SP_L(0))

	if len(self.MSAT_sect_2nd) == 0 {
		_ = binary.Write(&self.buf, binary.LittleEndian, SP_l(-2))
	} else {
		_ = binary.Write(&self.buf, binary.LittleEndian, SP_l(self.MSAT_sect_2nd[0]))
	}
	_ = binary.Write(&self.buf, binary.LittleEndian, SP_L(len(self.MSAT_sect_2nd)))
}

func (self *XlsDoc) Save(writer io.Writer, stream []byte) error {
	paddingLength := 0x1000 - (len(stream) % 0x1000)
	padding := make([]byte, paddingLength)

	self.book_stream_len = len(stream) + paddingLength

	self.BuildDirectory()
	self.BuildSat()

	self.WriteHeader()
	self.buf.Write(self.packed_MSAT_1st)
	self.buf.Write(stream)
	self.buf.Write(padding)
	self.buf.Write(self.packed_MSAT_2nd)
	self.buf.Write(self.packed_SAT)
	self.buf.Write(self.dir_stream)

	data := self.buf.Bytes()
	log.Println("save data:", len(data))
	_, err := writer.Write(data)
	return err
}
