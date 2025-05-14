package main

import (
	"bytes"
	"encoding/binary"
	"os"
)

const (
	SectorSize         = 512 // 标准扇区大小
	HeaderSize         = 512 // 文件头大小
	DirectoryEntrySize = 128 // 每个目录条目的大小
	MaxMiniStreamSize  = 4096
)

// OLE2 文件头结构
type OLE2Header struct {
	Signature            [8]byte  // 固定为 D0 CF 11 E0 A1 B1 1A E1
	CLSID                [16]byte // 类标识符（通常全0）
	MinorVersion         uint16
	MajorVersion         uint16
	ByteOrder            uint16 // 0xFFFE 表示小端
	SectorShift          uint16 // 扇区大小指数（如 9 表示 512 字节）
	MiniSectorShift      uint16 // 小扇区大小指数（通常 6 → 64 字节）
	Reserved             [6]byte
	NumDirectorySectors  uint32      // 目录扇区数
	NumFATSectors        uint32      // FAT 扇区数
	FirstDirectorySector uint32      // 第一个目录扇区位置
	MiniStreamCutoff     uint32      // 小流阈值（通常 4096）
	FirstMiniFATSector   uint32      // 第一个 MiniFAT 扇区
	NumMiniFATSectors    uint32      // MiniFAT 扇区数
	FirstDIFATSector     uint32      // 第一个 DIFAT 扇区
	NumDIFATSectors      uint32      // DIFAT 扇区数
	FAT                  [109]uint32 // 初始 FAT 表项
}

func NewOLE2Header() OLE2Header {
	header := OLE2Header{}
	copy(header.Signature[:], []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1})
	header.ByteOrder = 0xFFFE
	header.SectorShift = 9 // 512 字节扇区
	header.MiniSectorShift = 6
	header.MiniStreamCutoff = MaxMiniStreamSize
	header.MajorVersion = 3 // OLE2 版本
	return header
}

type DirectoryEntry struct {
	Name             [64]byte // UTF-16LE 编码的名称（最大31字符）
	NameLen          uint16
	EntryType        byte   // 0=空, 1=存储, 2=流, 5=根存储
	Color            byte   // 0=红, 1=黑（B树目录）
	LeftSibling      uint32 // 左兄弟条目ID
	RightSibling     uint32 // 右兄弟条目ID
	ChildSibling     uint32 // 子条目ID
	CLSID            [16]byte
	State            uint32
	CreationTime     uint64
	ModificationTime uint64
	StartSector      uint32 // 流起始扇区
	StreamSize       uint64 // 流大小（字节）
}

func CreateRootEntry() DirectoryEntry {
	entry := DirectoryEntry{}
	entry.EntryType = 5 // 根存储
	copy(entry.Name[:], []byte("Root Entry"))
	entry.NameLen = uint16(len("Root Entry") * 2) // UTF-16LE 长度
	return entry
}

func CreateStreamEntry(name string, startSector uint32, size uint64) DirectoryEntry {
	entry := DirectoryEntry{}
	entry.EntryType = 2 // 流
	nameUTF16 := bytesFromUTF16LE(name)
	copy(entry.Name[:], nameUTF16)
	entry.NameLen = uint16(len(nameUTF16))
	entry.StartSector = startSector
	entry.StreamSize = size
	return entry
}

// 辅助函数：将字符串转为 UTF-16LE 字节
func bytesFromUTF16LE(s string) []byte {
	buf := &bytes.Buffer{}
	for _, r := range s {
		binary.Write(buf, binary.LittleEndian, uint16(r))
	}
	return buf.Bytes()
}

func main() {
	// 1. 初始化头部
	header := NewOLE2Header()

	// 2. 创建目录条目（根目录和一个流）
	rootEntry := CreateRootEntry()
	streamEntry := CreateStreamEntry("ExampleStream", 0, 1024)

	// 3. 构建目录扇区
	dirSector := make([]byte, SectorSize)
	buf := bytes.NewBuffer(dirSector)
	binary.Write(buf, binary.LittleEndian, rootEntry)
	binary.Write(buf, binary.LittleEndian, streamEntry)

	// 4. 构建 FAT（此处简化，实际需处理链表）
	header.FAT[0] = 0xFFFFFFFE // 表示 FAT 结束

	// 5. 写入文件
	file, _ := os.Create("example.ole")
	defer file.Close()

	// 写入头部
	binary.Write(file, binary.LittleEndian, header)
	// 写入目录扇区
	file.Write(dirSector)
	// 写入其他扇区（如 FAT、数据流等，此处省略）
}
