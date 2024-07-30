def process_data(inputs):
    i1, i2, i3, i4, i5, i6, i7, i8, i9, i10 = inputs
    
    o1 = i8 - i7
    o2 = i10 - i9
    o3 = i5 - i6
    o4 = o1 / (i3 - i4) if (i3 - i4) != 0 else None
    o5 = o1 / o2 if o2 != 0 else None
    o6 = o4 / i1 if i1 != 0 else None
    
    return [o1, o2, o3, o4, o5, o6]
