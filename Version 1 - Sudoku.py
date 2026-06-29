"""
Sudoku solver
=============

Reads the GIVEN (starting) numbers of a Sudoku puzzle from an Excel table,
builds a full 81-cell model of the grid, and solves it using two classic
constraint techniques:

  1. Constraint propagation  (update_ranges):
        When a cell's value is known, remove that value from the candidate
        list ("RANGE") of every other cell that shares its row, column, or box.

  2. Hidden singles          (find_values):
        Within any row / column / box, if a digit can legally go in exactly
        one cell, that cell must hold that digit.

The two steps run in a loop until the grid is full, or until a pass finds
nothing new -- which means the puzzle needs a technique this solver does not
implement yet (see notes at the bottom).

Excel input
-----------
Worksheet name : "Range_to_numbers"
Row 1          : headers (skipped)
Columns (in order), one row per cell:
    A  INDIVIDUAL_CELL   cell name, e.g. "A1"  (column letter + row number)
    B  Value             the given digit 1-9, or 0 / blank / "?" if empty
    C  COLUMN_POSITION   1-9
    D  ROW_POSITION      1-9
    E  BOX_POSITION      1-9

NOTE: store the Value column as real numbers, not text. The matching logic
compares integers, so a value typed as the text "5" will not equal the
integer 5 and the solve will be wrong.
"""
# Initialize Sudoku cells as records
import openpyxl
from pathlib import Path

# ----------------------------------------------------------------------------
# 1. Load the given numbers from Excel
# ----------------------------------------------------------------------------

# Look for the workbook in the same folder as this script.
file_path = Path(__file__).parent / "Sudoku puzzle data.xlsx"

# data_only=True returns the *computed* cell values rather than formulas,
# so a cell containing "=5" comes back as 5.
wb = openpyxl.load_workbook(file_path, data_only=True)

# The worksheet that holds the one-row-per-cell table.
ws = wb["Range_to_numbers"]

# Build a list of records for the GIVEN cells only (the puzzle's clues).
given_flat_list = []
for row in ws.iter_rows(min_row=2, values_only=True):
    # Skip any cell that isn't actually filled in: 0, "0", None, blank, or "?".
    if row[1] in (0, "0", None, "", "?"):
        continue
    # One dict per given clue. The row[] indexes follow the column order above.
    record = {
        "INDIVIDUAL_CELL": row[0],       # Cell name, e.g. "A1"
        "Value": row[1],           # The given sudoku digit
        "COLUMN_POSITION": row[2],    # Column number 1-9
        "ROW_POSITION": row[3],    # Row number 1-9
        "BOX_POSITION": row[4]     # Box number 1-9
    }
    given_flat_list.append(record)


# ----------------------------------------------------------------------------
# 2. Build the full 81-cell template (every cell, given or not)
# ----------------------------------------------------------------------------

# This is the working model the solver mutates as it runs. Every cell starts
# blank with all nine digits as candidates; the givens get applied in step 7.
Sudoku_template_and_records = []
columns = "ABCDEFGHI"

# Outer loop = columns A-I (col_position 1-9); inner loop = rows 1-9.
for col_position, col_letter in enumerate(columns, start=1):
    for row_position in range(1, 10):
        
        # Work out which 3x3 box (1-9) this cell belongs to.
        #   (row-1)//3  -> row band 0,1,2 ; *3 spaces the bands out to 0,3,6
        #   (col-1)//3  -> column stack 0,1,2
        #   +1          -> shift from 0-indexed to boxes numbered 1-9
        # Boxes are numbered left-to-right, top-to-bottom (1,2,3 / 4,5,6 / 7,8,9).
        box_position = (
            ((row_position - 1) // 3) * 3
            + ((col_position - 1) // 3)
            + 1
        )

        sudoku_record = {
            "INDIVIDUAL_CELL": f"{col_letter}{row_position}",   # e.g. "A1"
            "Value": "",                                        # blank until solved
            "COLUMN_POSITION": col_position,                    # 1-9
            "ROW_POSITION": row_position,                       # 1-9
            "BOX_POSITION": box_position,                       # 1-9
            "RANGE_COUNT": 9,                                   # how many candidates remain
            "RANGE": list(range(1, 10))                        # candidate digits [1..9]
        }
        Sudoku_template_and_records.append(sudoku_record)
        
# ----------------------------------------------------------------------------
# 3. Constraint propagation: apply known values and prune their peers
# ----------------------------------------------------------------------------

def update_ranges(all_range, cells_given):
    """Apply every known value in `cells_given` to the full grid `all_range`.

    For each known cell:
      * collapse the cell itself to its single known value, and
      * remove that value from the candidate RANGE of every peer it shares a
        row, column, or box with.

    `cells_given` is reused for two purposes: the puzzle's original clues, and
    cells the solver discovers later. Mutates `all_range` in place; returns it.
    """
    
    for given in cells_given:
        temp_cell = given["INDIVIDUAL_CELL"]
        temp_col = given["COLUMN_POSITION"]
        temp_row = given["ROW_POSITION"]
        temp_box = given["BOX_POSITION"]
        temp_value = given["Value"]

        for picked_cell in all_range:
            if picked_cell["INDIVIDUAL_CELL"] == temp_cell:
                # This is the known cell itself -> lock it to its value.
                picked_cell["Value"] = temp_value
                picked_cell["RANGE"] = {temp_value}         # one candidate left (a set here)
            elif (picked_cell["RANGE_COUNT"] > 1 and
                  (picked_cell["COLUMN_POSITION"] == temp_col or
                   picked_cell["ROW_POSITION"] == temp_row or
                   picked_cell["BOX_POSITION"] == temp_box) and picked_cell["Value"] == '' and picked_cell["INDIVIDUAL_CELL"] != temp_cell):
                # A still-unsolved peer (same row, column, or box):
                # this value can no longer go here, so drop it from its candidates.
                if temp_value in picked_cell["RANGE"]:
                    picked_cell["RANGE"].remove(temp_value)
                
                # Keep the candidate count in sync with the candidate list.
                picked_cell["RANGE_COUNT"] = len(picked_cell["RANGE"])
    return all_range

# ----------------------------------------------------------------------------
# 4. Helpers: group the 81 cells into the 27 units (9 rows, 9 cols, 9 boxes)
# ----------------------------------------------------------------------------

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

# ----------------------------------------------------------------------------
# 5. Hidden singles: place a digit when only one cell in a unit can hold it
# ----------------------------------------------------------------------------


def find_values(all_range):
    """Scan every unit for "hidden singles" and fill them in.

    A hidden single: within one row / column / box, some digit i is not yet
    placed and only one cell still lists i as a candidate -> that cell must be i.

    Returns the list of cells filled in this pass, so the caller can propagate
    their effects with update_ranges.
    """
    edited_cells = []

    for unit in all_units(all_range):  # each row, then each column, then each box
        for i in range(1, 10):        # try every digit 1-9 in this unit

            counts_with_i = 0    # how many empty cells could still take i
            i_present = False    # is i already placed somewhere in this unit?
            picked_cell = None   # the lone candidate cell, if there is exactly one

            for cell in unit:

                if cell["Value"] == i:
                    # i is already used in this unit -> nothing to place.
                    i_present = True

                elif cell["Value"] == "" and i in cell["RANGE"]:
                    # An empty cell that can still take i.
                    picked_cell = cell
                    counts_with_i += 1

            # i isn't placed yet AND exactly one cell can take it -> place it there.
            if not i_present and counts_with_i == 1:
                picked_cell["Value"] = i
                picked_cell["RANGE"] = [i]
                picked_cell["RANGE_COUNT"] = 1
                edited_cells.append(picked_cell)

    return edited_cells

# ----------------------------------------------------------------------------
# 6. Small utilities for the main loop and output
# ----------------------------------------------------------------------------

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

# ----------------------------------------------------------------------------
# 7. Solve: seed the givens, then loop (find hidden singles -> propagate)
# ----------------------------------------------------------------------------

# Seed the grid with the puzzle's clues and prune their peers once up front.
update_ranges(Sudoku_template_and_records, given_flat_list)

k=0     # counts how many solving passes we make
while blanks_left(Sudoku_template_and_records) == True:
    # Find every hidden single currently on the board.
    returned_values = find_values(Sudoku_template_and_records)

    # If a whole pass found nothing new, the solver is stuck: this puzzle needs
    # a technique beyond hidden singles, so stop rather than loop forever.
    if len(returned_values) == 0:
        print("Solver stopped because no more values could be found.")
        break
    # Apply the newly found values and prune their peers, then loop again.
    update_ranges(Sudoku_template_and_records, returned_values)
    k=k+1

# ----------------------------------------------------------------------------
# 8. Results
# ----------------------------------------------------------------------------


print()
print("Final Sudoku Output:")
print_sudoku(Sudoku_template_and_records)



print()
print(f"Number of passes: {k}")
