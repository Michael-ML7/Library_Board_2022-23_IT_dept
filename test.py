def col_int_to_word(n):
    convertString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    base = 26
    i = n - 1

    if i < base:
        return convertString[i]
    else:
        return col_int_to_word(i // base) + convertString[i % base]