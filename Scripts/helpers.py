import re

# Take in a test name and remove anything in parentheses, EXCEPT if those parentheses contain "C/", which indicates temp/humidity info
def strip_parens_from_name(test):
    return re.sub(r"\((?![^()]*C/).+?\)", "", test)

# Return number of cells in a row and number of rows in a box depending on cell size
def box_size_mapper(cell_size):
    if cell_size in ['AA', 'AAA']:
        row_capacity = 11
        rows_in_box = 16
    elif cell_size == 'D':
        row_capacity = 5
        rows_in_box = 7
    elif cell_size == 'C':
        row_capacity = 7
        rows_in_box = 10
    else: #in case something weird happens and it doesn't match, input dummy values then change DF after for loop
        row_capcity = 1 
        rows_in_box = 1
    return row_capacity, rows_in_box


# Take in list of cell ranges from dataframe, number of lots, and the cell size -> return number of boxes needed for
# each group, stored in a list. 
def get_number_of_boxes(list_of_cell_ranges, number_of_lots, cell_size):
    row_capacity, rows_in_box = box_size_mapper(cell_size)
    boxes_per_group = []
    for cell_range in list_of_cell_ranges:
        if isinstance(cell_range, str):
            # Get number of cells per test group
            cells = cell_range.split('-')
            number_of_cells = int(cells[1]) - int(cells[0]) + 1
            # Get number of rows needed per test group
            number_of_rows_per_lot = number_of_cells // row_capacity
            if (number_of_cells % row_capacity) > 0:      # if cells don't complete a full row
                number_of_rows_per_lot += 1
            total_number_of_rows = int(number_of_rows_per_lot) * int(number_of_lots)
            # Get number of boxes needed per test group
            number_of_boxes = total_number_of_rows // rows_in_box
            if (total_number_of_rows % rows_in_box) > 0:  # if rows don't complete a full box
                number_of_boxes += 1
            boxes_per_group.append(number_of_boxes)  
    return boxes_per_group


# Format lot numbers
# Takes list of lot numbers and returns ranges
def format_lots(numbers):
    formatted_list = []
    start = numbers[0]   #keep track of range starts
    end = start          #keep track of range ends

    for n in numbers[1:]:
        if n == end + 1:   #if consecutive
            end = n
        else:
            if start == end:   #single lot
                formatted_list.append(str(start))
            else:              #lot range
                formatted_list.append(f"{start}-{end}")
            # Start a new range
            start = n
            end = n
    # Add the last range or number
    if start == end:
        formatted_list.append(str(start))
    else:
        formatted_list.append(f"{start}-{end}")
    return ', '.join(formatted_list)
