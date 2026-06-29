# Initialize Sudoku cells as records
import openpyxl
from pathlib import Path

# The Excel file should be in the same folder as your Python script
file_path = Path(__file__).parent / "Sudoku puzzle data.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
# Sheet with your table
ws = wb["Range_to_numbers"]

#Get flat list of everything that I have for the Excel file
given_flat_list = []
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[1] in (0, "0", None, "", "?"):
        continue
    record = {
        "INDIVIDUAL_CELL": row[0],       # A1, A2, etc.
        "Value": row[1],           # Sudoku value
        "COLUMN_POSITION": row[2],    # Column number
        "ROW_POSITION": row[3],    # Row number
        "BOX_POSITION": row[4]     # Box number
    }
    given_flat_list.append(record)

#Create the FULL template
Sudoku_template_and_records = []
columns = "ABCDEFGHI"
for col_position, col_letter in enumerate(columns, start=1):
    for row_position in range(1, 10):

        box_position = (
            ((row_position - 1) // 3) * 3
            + ((col_position - 1) // 3)
            + 1
        )

        sudoku_record = {
            "INDIVIDUAL_CELL": f"{col_letter}{row_position}",
            "Value": "",
            "COLUMN_POSITION": col_position,
            "ROW_POSITION": row_position,
            "BOX_POSITION": box_position,
            "RANGE_COUNT": 9,
            "RANGE": list(range(1, 10))
        }
        Sudoku_template_and_records.append(sudoku_record)

#Function to update ranges as need be
def update_ranges(all_range, cells_given):
    """Apply Function #1.  Mutates all_range in place; returns it."""
    for given in cells_given:
        temp_cell = given["INDIVIDUAL_CELL"]
        temp_col = given["COLUMN_POSITION"]
        temp_row = given["ROW_POSITION"]
        temp_box = given["BOX_POSITION"]
        temp_value = given["Value"]

        for picked_cell in all_range:
            if picked_cell["INDIVIDUAL_CELL"] == temp_cell:
                # the given cell itself -> collapse it to the known value
                picked_cell["Value"] = temp_value
                picked_cell["RANGE"] = {temp_value}        # rangecount becomes 1
            elif (picked_cell["RANGE_COUNT"] > 1 and
                  (picked_cell["COLUMN_POSITION"] == temp_col or
                   picked_cell["ROW_POSITION"] == temp_row or
                   picked_cell["BOX_POSITION"] == temp_box) and picked_cell["Value"] == '' and picked_cell["INDIVIDUAL_CELL"] != temp_cell):
                # a peer that still has options -> eliminate this value
                # remove temp_value only if it exists in the list
                if temp_value in picked_cell["RANGE"]:
                    picked_cell["RANGE"].remove(temp_value)

                picked_cell["RANGE_COUNT"] = len(picked_cell["RANGE"])
    return all_range


def get_rows(all_range):
    rows = []

    for row_num in range(1, 10):
        one_row = []

        for cell in all_range:
            if cell["ROW_POSITION"] == row_num:
                one_row.append(cell)

        rows.append(one_row)

    return rows


def get_columns(all_range):
    columns = []

    for col_num in range(1, 10):
        one_column = []

        for cell in all_range:
            if cell["COLUMN_POSITION"] == col_num:
                one_column.append(cell)

        columns.append(one_column)

    return columns


def get_boxes(all_range):
    boxes = []

    for box_num in range(1, 10):
        one_box = []

        for cell in all_range:
            if cell["BOX_POSITION"] == box_num:
                one_box.append(cell)

        boxes.append(one_box)

    return boxes

def all_units(all_range):
    Rows = get_rows(all_range)
    Columns = get_columns(all_range)
    Boxes = get_boxes(all_range)

    return Rows + Columns + Boxes

def find_values(all_range):
    edited_cells = []

    for unit in all_units(all_range):  # each row, each column, each box
        for i in range(1, 10):

            counts_with_i = 0
            i_present = False
            picked_cell = None

            for cell in unit:

                if cell["Value"] == i:
                    i_present = True

                elif cell["Value"] == "" and i in cell["RANGE"]:
                    picked_cell = cell
                    counts_with_i += 1

            if not i_present and counts_with_i == 1:
                picked_cell["Value"] = i
                picked_cell["RANGE"] = [i]
                picked_cell["RANGE_COUNT"] = 1
                edited_cells.append(picked_cell)

    return edited_cells

def blanks_left(all_range):
    """Returns True if there are any blank Sudoku cells left."""
    for cell in all_range:
        if cell["Value"] == "":
            return True

    return False

def print_sudoku(all_range):
    """Prints the final Sudoku grid."""

    for row_num in range(1, 10):
        row_values = []

        for col_num in range(1, 10):
            for cell in all_range:
                if cell["ROW_POSITION"] == row_num and cell["COLUMN_POSITION"] == col_num:
                    value = cell["Value"]

                    if value == "":
                        row_values.append(".")
                    else:
                        row_values.append(str(value))

        print(" ".join(row_values))

update_ranges(Sudoku_template_and_records, given_flat_list)

k=0
while blanks_left(Sudoku_template_and_records) == True:
    returned_values = find_values(Sudoku_template_and_records)

    # Prevent infinite loop if no new values are found
    if len(returned_values) == 0:
        print("Solver stopped because no more values could be found.")
        break

    update_ranges(Sudoku_template_and_records, returned_values)
    k=k+1

print()
print("Final Sudoku Output:")
print_sudoku(Sudoku_template_and_records)



print()
print(f"Number of passes: {k}")
