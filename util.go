package xlwt

func FillBytes(size int, value byte) []byte {
	buf := make([]byte, size)
	for i := 0; i < size; i++ {
		buf[i] = value
	}
	return buf
}

func FillInt(size int, value int) []int {
	buf := make([]int, size)
	for i := 0; i < size; i++ {
		buf[i] = value
	}
	return buf
}
