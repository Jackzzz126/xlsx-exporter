def pos_index_2_str(row_index, col_index):
	col = col_index + 1
	colStr = ""

	while True:
		if col % 26 > 0:
			colStr = chr(ord("A") + col % 26 - 1) + colStr
			col = col / 26
		else:#col % 26 = 0
			if col == 0:
				break
			else:
				colStr = chr(ord("A") + 26 - 1) + colStr
				col = col / 26 - 1
	return colStr + str(row_index + 1)
