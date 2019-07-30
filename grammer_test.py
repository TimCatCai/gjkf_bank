def a(size):
    count = 0

    def b():
        nonlocal count
        if count < size:
            count += 1
            return 1
        else:
            return 0
    return b



