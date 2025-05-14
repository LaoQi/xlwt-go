package xlwt

const SECTOR_SIZE = 0x0200
const MIN_LIMIT = 0x1000

const SID_FREE_SECTOR = -1
const SID_END_OF_CHAIN = -2
const SID_USED_BY_SAT = -3
const SID_USED_BY_MSAT = -4

type SP_L uint32 // struct.pack('<L')
type SP_l int32  // struct.pack('<l')
type SP_H uint16 // struct.pack('<H')
type SP_I uint32 // struct.pack('<I')
type SP_B uint8  // struct.pack('<B')

var SP_H_0 = []byte{0x00, 0x00}
var SP_H_1 = []byte{0x01, 0x00}
