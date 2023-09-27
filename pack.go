package xlsy

func packInt32sToInt64(a, b int32) int64 {
	// Convert int32 values to int64 and shift them into position.
	// Assuming a and b are both 32-bit integers.
	return (int64(a) << 32) | int64(b)
}

func unpackInt64ToInt32s(val int64) (int32, int32) {
	// Extract the two int32 values from the int64.
	// Right-shifting by 32 bits retrieves the first value,
	// and masking with 0xFFFFFFFF retrieves the second value.
	return int32(val >> 32), int32(val & 0xFFFFFFFF)
}
