import pyautogui
import mouse
import os

'''
3 different types of action: double click, single click and typing. Create 1 function that records these actions and
1 function that uses these coordinates for actions.

Saving coordinates:
    print instruction
    wait for lmb
    record current position using pyautogui.position() or other choice for just x,y values
    write coordinates to text file
    perform click or double click depending on action
    typewrite() if needed
'''

template_num = 1554  # Sage100 template number - unsure of what this is exactly, but it's required

cut_work_order_page_open_time = 5  # Time in seconds to open page to cut/open work orders


def save_coordinate(filepath: str, instruction: str, dbl_clk: bool, typewrite: bool, typewrite_val: str):
    """
       Saves the current mouse coordinates to a file and performs a mouse click and optional typing action.

       Args:
           instruction (str): Instruction or message to print before saving the coordinates.
           dbl_clk (bool): Whether to perform a double-click. Default is False.
           typewrite (bool): Whether to type a string after clicking. Default is False.
           typewrite_val (str): The string to type if typewrite is True. Default is an empty string.
           filepath (str): Filepath to save coordinates for portion being trained
       """

    print(instruction)
    mouse.wait('left')
    x, y = pyautogui.position()

    coords = f"{x},{y}\n"

    write_file(filepath=filepath, thing_to_write=coords)

    if dbl_clk:
        pyautogui.doubleClick()

    else:
        pyautogui.click()

    if typewrite and typewrite_val:
        pyautogui.typewrite(typewrite_val)

    return


def write_file(filepath: str, thing_to_write: str):
    if os.path.exists(filepath):

        with open(filepath, "a") as file:
            file.write(thing_to_write)

        file.close()

    else:
        with open(filepath, "w") as file:
            file.write(thing_to_write)

        file.close()


# Main training loop - will record all coordinates and cut 1 work order
def main_train_loop(pn: str, qty: str, job_num: str, instructions: list, step_count: int):
    # Position is the order in which actions are performed, starting from 0
    # For example, a double click happens at steps 13 and 16 in the automation process
    dbl_clk_pos = [13, 16]
    typewrite_pos = [2, 3, 4, 5]

    # Loop through each instruction in the list to save each coordinate accordingly.
    for item in instructions:
        dbl_clk = False
        typewrite = False
        type_val = None

        if step_count in dbl_clk_pos:
            dbl_clk = True

        if step_count in typewrite_pos:
            typewrite = True

            # Typewrite either PN, QTY, Job Number or template number into boxes, depending on position in steps

            if step_count == 2:
                type_val = pn

            elif step_count == 3:
                type_val = qty

            elif step_count == 4:
                type_val = job_num

            else:
                type_val = str(template_num)

        save_coordinate(filepath='cut_wo_coords.txt', instruction=item, dbl_clk=dbl_clk, typewrite=typewrite,
                        typewrite_val=type_val)
        step_count += 1


def main_train_loop_releasing_completing_closing(wo_num: str, qty: str, instructions: list):
    """release_complete_close_wo_instructions = [
        'Recording "Work Ticket No." text box location...',
        'Recording "Status" dropdown box location...',
        'Recording "Released" dropdown box selection location...',
        'Recording "Accept" dropdown button location...',
        'Recording X Button on Print Material Shortages Screen...',
        'Recording "Work Ticket Entry" window X button...',
        'Recording "Work Ticket Transaction Register/Update button...',
        'Recording "Print" button...',
        'Recording "Work Ticket Transaction Entry" button...',
        'Recording "Next transaction #" button...',
        'Recording Tab 2 button location...',
        'Recording "Work Ticket No." box location...
        '

    ]"""

    # This section does the WO releasing portion of the automation task.
    # This assumes that the training for cutting was completed directly prior to this training

    # WO Ticket Number box
    print(instructions[0])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()
    pyautogui.typewrite(wo_num)
    pyautogui.press('Enter')

    # Status dropdown box
    print(instructions[1])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.doubleClick()

    # "Released" dropdown option
    print(instructions[2])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # "Yes" / "Accept" button to confirm releasing
    print(instructions[3])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # "X" to close out of print window
    print(instructions[4])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # "X" to close out of work ticket entry window
    print(instructions[5])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # Work Ticket Transaction Register/Update
    print(instructions[6])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # Print button
    print(instructions[7])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # This section does the WO completing portion of the automation task.

    # Work Ticket transaction entry
    print(instructions[8])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # Next Trans #
    print(instructions[9])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # Tab 2 - Lines
    print(instructions[10])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # WO Number Box
    print(instructions[11])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.doubleClick()
    pyautogui.typewrite(wo_num)

    # WO Qty Box
    print(instructions[12])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.doubleClick()
    pyautogui.typewrite(qty)

    # Accept
    print(instructions[13])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # Print
    print(instructions[14])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # Print
    print(instructions[15])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # Window Exit
    print(instructions[16])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # Yes Button
    print(instructions[17])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # This section does the WO closing portion of the automation task.

    # Work Ticket Trans. Entry Button
    print(instructions[18])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # Next Trans. Number button
    print(instructions[19])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # Status dropdown box
    print(instructions[20])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.doubleClick()

    # "Closed" dropdown box selection
    print(instructions[21])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # Tab 2 - Lines
    print(instructions[22])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    '''
    # Tab 2 - Lines
    print(instructions[23])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()
    '''

    # WO Number Box
    print(instructions[23])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.doubleClick()
    pyautogui.typewrite(wo_num)

    # Accept
    print(instructions[24])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    # Print
    print(instructions[25])
    mouse.wait('left')
    x, y = pyautogui.position()
    coords = f"{x},{y}\n"
    write_file('release_complete_close_wo_coords.txt', coords)
    pyautogui.click()

    return

# main_train_loop(pn="3768-2", qty=10, job_num='3016-023', instructions=wo_cut_instructions)
