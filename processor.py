import statistics

def process_data(inputs):
    i1, i2, i3, i4, i5, i6, i7, i8, i9, i10 = inputs
    
    o1 = i8 - i7
    o2 = i10 - i9
    o3 = i5 - i6
    o4 = o1 / (i3 - i4) if (i3 - i4) != 0 else None
    o5 = o1 / o2 if o2 != 0 else None
    o6 = o4 / i1 if i1 != 0 and o4 is not None else None
    
    return [o1, o2, o3, o4, o5, o6]

def calculate_additional_outputs(results, base_o9, base_o12, margin_of_error):
    o5_values = [result[4] for result in results if result[4] is not None]
    o6_values = [result[5] for result in results if result[5] is not None]

    o7 = statistics.mean(o5_values) if o5_values else None
    o8 = statistics.stdev(o5_values) if len(o5_values) > 1 else None
    o9 = o8 / o7 if o7 and o8 else None
    o10 = statistics.mean(o6_values) if o6_values else None
    o11 = statistics.stdev(o6_values) if len(o6_values) > 1 else None
    o12 = o11 / o10 if o10 and o11 else None

    warning_o9 = abs(o9 - base_o9) > margin_of_error if o9 is not None else False
    warning_o12 = abs(o12 - base_o12) > margin_of_error if o12 is not None else False

    return [o7, o8, o9, o10, o11, o12], warning_o9, warning_o12
