import tkinter as tk
from tkinter import messagebox
import time
import pyautogui
import mouse
import os
import re


def strings_to_tuples(strings_list):
    """
    Turns a list of strings of coordinates 'x,y' into a list of integer tuples (x,y)
    :param strings_list:
    :return:
    """
    tuples_list = []

    for string in strings_list:
        x_str, y_str = string.split(',')
        x_int = int(x_str)
        y_int = int(y_str)
        tuples_list.append((x_int, y_int))

    return tuples_list


class MyApp:
    def __init__(self, root):
        self.cut_wo_instructions = [
            'Recording "Work Ticket Entry" button location...',
            'Recording "Next Work Order Number" button location...',
            'Recording "Item Code" box location...',
            'Recording "Quantity Ordered" box location...',
            'Recording "Job Number" box location...',
            'Recording "Template Number" box location...',
            'Recording "Steps" button location...',
            'Recording "Materials" button location...',
            'Recording "Inquiry" button location...',
            'Recording "Schedule" button location...',
            'Recording "Totals" button location...',
            'Recording "Accept" button location...',
        ]

        self.release_complete_close_wo_instructions = [
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
            'Recording Tab 2 - Lines" button...',
            'Recording WO Number Box location...',
            'Recording WO Quantity location...',
            'Recording "Accept" button location...',
            'Recording "Print" button location...',
            'Recording "Print" button location...',
            'Recording Window X button location...',
            'Recording "Yes" button location...',
            'Recording "Work Ticket Transaction Entry" button location...',
            'Recording "Next Transaction Number " button location...',
            'Recording "Status" dropdown box location...',
            'Recording "Closed" dropdown box selection location...',
            'Recording Tab 2 - Lines" button...',
            'Recording "Work Ticket No." text box location...',
            'Recording "Accept" button location...',
            'Recording "Print" button location...',
        ]

        self.cut_work_order_page_open_time = 5  # Time in seconds to open page to cut/open work orders
        self.template_num = 1554

        self.root = root
        self.root.geometry("1000x550")  # Set the window size
        self.root.resizable(False, False)
        self.root.iconbitmap('RPG Favicon.ico')
        self.root.title("RPG Work Order Cutter/Completer/Closer")

        # Create a frame to hold all the widgets
        self.frame = tk.Frame(self.root)
        self.frame.pack(expand=True)

        # Create labels and text boxes for Part Number, Quantity, Job Number, and Work Order Number
        self.labels = ['Part Number', 'Quantity', 'Job Number', 'Work Order Number']
        self.entries = []

        # Create an entry box for each label defined above
        for i, label in enumerate(self.labels):
            tk.Label(self.frame, text=label, font="Segoe_UI 12 underline").grid(row=0, column=i, padx=10, pady=5)
            entry = tk.Text(self.frame, width=20, height=20)  # Define the size, padding of the boxes (20 lines)
            entry.grid(row=1, column=i, padx=10, pady=5)  # Place each box iteratively
            self.entries.append(entry)  # Keep track of the entry boxes after their initialized

        # Create the menu bar
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        # Create 'Actions' menu
        self.actions_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Menu", menu=self.actions_menu)

        self.actions_menu.add_command(label="Instructions", command=self.show_info_window)
        self.actions_menu.add_command(label="Train W/O Cutting", command=self.train_cutting)
        self.actions_menu.add_command(label="Execute W/O Cutting", command=self.execute_cutting_wos)
        self.actions_menu.add_command(label="Save Work Order Numbers", command=self.save_wo_nums)
        self.actions_menu.add_command(label="Train W/O Releasing & Completing & Closing",
                                      command=self.train_releasing_completing_closing)

        self.actions_menu.add_command(label="Exit", command=self.exit_program)

        # Create instructions text box and span all columns
        self.instructions_txtbox = tk.Text(self.frame, height=2)
        self.instructions_txtbox.insert(tk.END, "Training instructions will appear here.")
        self.instructions_txtbox.configure(state="disabled")
        self.instructions_txtbox.grid(row=3, column=0, columnspan=4, sticky='ew', padx=10, pady=5)

        self.execute_button = tk.Button(text="Execute Cutting W/Os", command=self.execute_cutting_wos)
        self.execute_button.pack(pady=30)

        # Regex for matching part numbers: [0-9][0-9][0-9][0-9]-[0-9][0-9]?
        # Regex for matching UYZ08- part numbers: ^UYZ08-.+

    def show_info_window(self):
        messagebox.showinfo("Welcome - Instructions",
                            "Welcome to the RPG Work Order Cutter/Completer/Closer application!\n\n"
                            "Here are the instructions:\n\n"
                            "- To train the application for cutting work orders, use the 'Train W/O Cutting' "
                            "option.\n\n"
                            "- Training consists of following the displayed instructions and clicking when "
                            "instructed.\n\n"
                            "- There must be only one (1) entry for the training process. The training process will "
                            "cut one (1) work order.\n\n"
                            "- To execute the work order cutting process, use the 'Execute W/O Cutting' option.\n\n"
                            "- There must be at least one (1) entry in the PN, Qty and Job# boxes and the training "
                            "must be completed to execute.\n\n"
                            "- For saving work order numbers, use the 'Save Work Order Numbers' option.\n\n"
                            "- Any Work Order Numbers will be saved to 'wo_nums.txt' in the same directory that this "
                            "executable is found in.\n\n"
                            "- To exit the application, use the 'Exit' option or press the red 'X' button.\n\n")

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

    def save_coordinate(self, filepath: str, instruction: str, dbl_clk: bool, typewrite: bool, typewrite_val: str):
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

        self.write_file(filepath=filepath, thing_to_write=coords)

        if dbl_clk:
            pyautogui.doubleClick()

        else:
            pyautogui.click()

        if typewrite and typewrite_val:
            pyautogui.typewrite(typewrite_val)

        return

    def write_file(self, filepath: str, thing_to_write: str):
        """
        If a file does not exist, create the file and write to it. Otherwise, append to the existing file
        :param filepath: string that dictates filepath of file being written to
        :param thing_to_write: thing that will be written into the file
        :return:
        """

        if os.path.exists(filepath):

            with open(filepath, "a") as file:
                file.write(thing_to_write)

            file.close()

        else:
            with open(filepath, "w") as file:
                file.write(thing_to_write)

            file.close()

    # Main training loop - will record all coordinates and cut 1 work order
    def main_train_loop(self, pn: str, qty: str, job_num: str, instructions: list, step_count: int):
        # Position is the order in which actions are performed, starting from 0
        # For example, a double click happens at steps 13 and 16 in the automation process
        dbl_clk_pos = [13, 16]
        typewrite_pos = [2, 3, 4, 5]

        # Loop through each instruction in the list to save each coordinate accordingly.
        for item in instructions:

            # self.instructions_label['text'] = item
            self.instructions_txtbox.configure(state="normal")
            self.instructions_txtbox.delete("1.0", "end")
            self.instructions_txtbox.insert(tk.END, item)
            self.instructions_txtbox.update()
            self.instructions_txtbox.configure(state="disabled")
            # self.instructions_txtbox.update_idletasks()  # Force the UI to refresh

            dbl_clk = False
            typewrite = False
            type_val = None

            # Step after "next work order number" button is pressed
            if step_count == 2:
                self.prompt_work_order_number()

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
                    type_val = str(self.template_num)

            self.save_coordinate(filepath='cut_wo_coords.txt', instruction=item, dbl_clk=dbl_clk, typewrite=typewrite,
                                 typewrite_val=type_val)
            step_count += 1

    def main_train_loop_releasing_completing_closing(self, wo_num: str, qty: str, instructions: list):

        # This section does the WO releasing portion of the automation task.
        # This assumes that the training for cutting was completed directly prior to this training
        # Requires quantity and WO number to function

        # WO Ticket Number box
        print(instructions[0])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()
        pyautogui.typewrite(wo_num)
        pyautogui.press('Enter')

        # Status dropdown box
        print(instructions[1])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.doubleClick()

        # "Released" dropdown option
        print(instructions[2])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # "Yes" / "Accept" button to confirm releasing
        print(instructions[3])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # "X" to close out of print window
        print(instructions[4])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # "X" to close out of work ticket entry window
        print(instructions[5])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # Work Ticket Transaction Register/Update
        print(instructions[6])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # Print button
        print(instructions[7])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # This section does the WO completing portion of the automation task.

        # Work Ticket transaction entry
        print(instructions[8])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # Next Trans #
        print(instructions[9])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # Tab 2 - Lines
        print(instructions[10])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # WO Number Box
        print(instructions[11])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.doubleClick()
        pyautogui.typewrite(wo_num)

        # WO Qty Box
        print(instructions[12])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.doubleClick()
        pyautogui.typewrite(qty)

        # Accept
        print(instructions[13])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # Print
        print(instructions[14])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # Print
        print(instructions[15])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # Window Exit
        print(instructions[16])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # Yes Button
        print(instructions[17])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # This section does the WO closing portion of the automation task.

        # Work Ticket Trans. Entry Button
        print(instructions[18])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # Next Trans. Number button
        print(instructions[19])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # Status dropdown box
        print(instructions[20])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.doubleClick()

        # "Closed" dropdown box selection
        print(instructions[21])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # Tab 2 - Lines
        print(instructions[22])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
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
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.doubleClick()
        pyautogui.typewrite(wo_num)

        # Accept
        print(instructions[24])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        # Print
        print(instructions[25])
        mouse.wait('left')
        x, y = pyautogui.position()
        coords = f"{x},{y}\n"
        self.write_file('release_complete_close_wo_coords.txt', coords)
        pyautogui.click()

        return

    def read_values(self):
        """
        Reads all the boxes
        :return: values, or list all the things typed in the all the boxes
        """
        values = [entry.get("1.0", "end-1c").strip() for entry in self.entries]
        # print(values)  # To verify the values read from the text boxes
        return values

    def train_cutting(self):

        """
        Function that is linked to train w/o cutting button. Prompts user to verify that they'd like to train cutting
        coordinates. If yes, wipe current file and continue with training process. If no, return and do nothing.

        :return: None
        """

        answer = messagebox.askyesno("Confirm Training",
                                     "Are you sure you want to train? This will wipe the existing file.")

        if answer:

            # Wipe the file once training begins
            if os.path.isfile("cut_wo_coords.txt"):
                with open("cut_wo_coords.txt", "w") as file:
                    file.close()

            box_vals = self.read_values()

            pns = box_vals[0].split("\n")
            qtys = box_vals[1].split("\n")
            job_nos = box_vals[2].split("\n")

            if not len(pns) == len(qtys) == len(job_nos):
                self.show_error(message="Number of part numbers does not match number of quantities or job numbers!")
                return

            match1 = re.compile("^[0-9]{4}-[0-9]{1,2}$")
            match2 = re.compile("^UYZ08-.+$")
            match3 = re.compile("^[+-]?[0-9]+$")
            match4 = re.compile("^[0-9]{4}-[0-9]{3}$")

            for index in range(len(pns)):
                res1 = match1.match(pns[index])
                res2 = match2.match(pns[index])
                res3 = match3.match(qtys[index])
                res4 = match4.match(job_nos[index])

                if not res1 and not res2:
                    self.show_error(message="Double check formatting for part numbers! Incorrect characters detected.")
                    return

                elif not res3:
                    self.show_error(message="Double check formatting for quantities! Incorrect characters detected.")
                    return

                elif not res4:
                    self.show_error(message="Double check formatting for job numbers! Incorrect characters detected.")
                    return

            # Check if the first three boxes are empty or contain more than one object
            if not box_vals[0] or not box_vals[1] or not box_vals[2]:
                self.show_error("Error: One or more required fields are empty.")
            elif len(pns) > 1 or len(qtys) > 1 or len(job_nos) > 1:
                self.show_error("Error: More than one object in one of the three required fields.")
            else:
                print("Proceeding with training...")
                self.main_train_loop(pn=pns[0], qty=qtys[0], job_num=job_nos[0],
                                     instructions=self.cut_wo_instructions, step_count=0)

            print(box_vals)

        else:
            return

    def show_error(self, message):
        # Displays tkitner error box with custom message inside it - dissapears after 2 seconds
        error_box = tk.Toplevel(self.root)
        error_box.title("Error")
        tk.Label(error_box, text=message).pack(padx=20, pady=20)
        self.root.after(2000, error_box.destroy)  # Destroy the error box after 2000 milliseconds (2 seconds)

    def prompt_work_order_number(self):
        # Creates a tertiary pop-up window that prompts the user to enter a work order #
        # Submit button is linked to append_work_order_number()
        self.tertiary_window = tk.Toplevel(self.root)
        self.tertiary_window.geometry("200x200")
        self.tertiary_window.title("Enter Work Order Number")
        self.tertiary_window.attributes("-topmost", True)
        self.tertiary_window.iconbitmap('RPG Favicon.ico')

        self.tertiary_window.lift()
        self.tertiary_window.focus_force()

        tk.Label(self.tertiary_window, text="Work Order Number:").pack(padx=20, pady=10)
        self.wo_entry = tk.Entry(self.tertiary_window)
        self.wo_entry.focus_set()
        self.wo_entry.pack(padx=20, pady=10)

        tk.Button(self.tertiary_window, text="Submit", command=self.append_work_order_number).pack(padx=20, pady=10)

        self.root.wait_window(self.tertiary_window)

    def append_work_order_number(self):

        """
        Function that appends the entered work order number from the pop-up prompt to the textbox
        :return: None
        """
        wo_number = self.wo_entry.get().strip()
        if wo_number:
            current_text = self.entries[-1].get("1.0", "end-1c").strip()
            if current_text:
                new_text = current_text + "\n" + wo_number
            else:
                new_text = wo_number
            self.entries[-1].delete("1.0", "end")
            self.entries[-1].insert(tk.END, new_text)
        self.tertiary_window.destroy()

    def train_releasing_completing_closing(self):
        """
        Linked to the "train releasing/completing/closing WO" button.
        Prompts user to confirm training. If yes, wipes old coordinates/creates file and begins the training
        process for releasing/completing/closing. Checks for empty boxes and checks that there is only one WO/Qty combo.


        :return: None
        """
        # If the file exists, wipe it clean before training again
        answer = messagebox.askyesno("Confirm Training",
                                     "Are you sure you want to train? This will wipe the existing file.")

        if answer:

            if os.path.isfile("release_complete_close_wo_coords.txt"):
                with open("release_complete_close_wo_coords.txt", "w") as file:
                    file.close()

            box_vals = self.read_values()
            wo_nums = box_vals[3].split("\n")
            qtys = box_vals[1].split("\n")

            if not box_vals[3].strip():  # Corrected condition to detect empty box
                self.show_error("Error: Work Order Number box is empty.")

            if not box_vals[1].strip():
                self.show_error("Error: Quantity box is empty")

            elif len(wo_nums) > 1:
                self.show_error("Error: More than one object in Work Order Numbers box for training.")

            else:
                qty_train = qtys[0]
                wo_num_train = wo_nums[0]
                print("Proceeding with training...")
                self.main_train_loop_releasing_completing_closing(wo_num=wo_num_train,
                                                                  qty=qty_train,
                                                                  instructions=self.release_complete_close_wo_instructions)
        else:
            return

    def execute_cutting_wos(self):
        # Regex for matching part numbers: [0-9][0-9][0-9][0-9]-[0-9][0-9]? or ^\d{4}-\d{1,2}$
        # Regex for matching UYZ08- part numbers: ^UYZ08-.+$ ex.
        # Regex for matching quantities: ^[+-]?\d+$ ex. 1, 2, 3, 4...
        # Regex for matching job numbers: ^\d{4}-\d{3}$

        box_vals = self.read_values()

        pns = box_vals[0].split("\n")
        qtys = box_vals[1].split("\n")
        job_nos = box_vals[2].split("\n")

        # Check if the first three boxes are empty or contain more than one object
        if not box_vals[0] or not box_vals[1] or not box_vals[2]:
            self.show_error("Error: One or more required fields are empty.")
            return

        with open("cut_wo_coords.txt", "r") as file:
            lines = file.read().splitlines()

            if len(lines) is None:
                self.show_error("Error: Coordinates need to be trained before executing!")
                return

            if len(lines) != 12:
                self.show_error("Error: Incomplete training detected - retrain before executing!")
                return

            file.close()

        coord_list = strings_to_tuples(lines)

        if not len(pns) == len(qtys) == len(job_nos):
            self.show_error(message="Number of part numbers does not match number of quantities or job numbers!")
            return

        match1 = re.compile("^[0-9]{4}-[0-9]{1,2}$")
        match2 = re.compile("^UYZ08-.+$")
        match3 = re.compile("^[+]?[0-9]+$")
        match4 = re.compile("^[0-9]{4}-[0-9]{3}$")

        for index in range(len(pns)):
            res1 = match1.match(pns[index])
            res2 = match2.match(pns[index])
            res3 = match3.match(qtys[index])
            res4 = match4.match(job_nos[index])

            if not res1 and not res2:
                self.show_error(message="Double check formatting for part numbers! Incorrect characters detected.")
                return

            elif not res3:
                self.show_error(message="Double check formatting for quantities! Incorrect characters detected.")
                return

            elif not res4:
                self.show_error(message="Double check formatting for job numbers! Incorrect characters detected.")
                return

        for index in range(len(pns)):
            curr_pn = pns[index]
            curr_qty = qtys[index]
            curr_job_no = job_nos[index]

            self.cut_wo(pn=curr_pn, qty=curr_qty, job_no=curr_job_no, coordinates=coord_list)

    def cut_wo(self, pn: str, qty: str, job_no: str, coordinates: list):
        # 12 steps to cutting W/O
        # This is assuming that no other window is open on Sage and that the user is on the Production Mgmt
        WAIT_TIME_ENTRY = 5

        x_vals = [x for x, y in coordinates]
        y_vals = [y for x, y in coordinates]

        # First, open the work ticket entry window
        pyautogui.moveTo(x_vals[0], y_vals[0], 0.25)
        pyautogui.click()
        time.sleep(WAIT_TIME_ENTRY)

        # Click "Next Work Ticket Number" button
        pyautogui.moveTo(x_vals[1], y_vals[1], 0.25)
        pyautogui.click()

        # Prompt for the next work order number
        self.prompt_work_order_number()

        # Move to Item Code Box - typewrite item code
        pyautogui.moveTo(x_vals[2], y_vals[2], 0.25)
        pyautogui.click()
        pyautogui.typewrite(pn)

        # Move to quantity box - typewrite quantity
        pyautogui.moveTo(x_vals[3], y_vals[3], 0.25)
        pyautogui.click()
        pyautogui.typewrite(qty)

        # Move to job # box - typewrite job #
        pyautogui.moveTo(x_vals[4], y_vals[4], 0.25)
        pyautogui.click()
        pyautogui.typewrite(job_no)

        # Move to template # box - typewrite template #
        pyautogui.moveTo(x_vals[5], y_vals[5], 0.25)
        pyautogui.click()
        pyautogui.typewrite(str(self.template_num))

        # Steps tab
        pyautogui.moveTo(x_vals[6], y_vals[6], 0.25)
        pyautogui.click()

        # Materials tab
        pyautogui.moveTo(x_vals[7], y_vals[7], 0.25)
        pyautogui.click()

        # Inquiry tab
        pyautogui.moveTo(x_vals[8], y_vals[8], 0.25)
        pyautogui.click()

        # Schedule tab
        pyautogui.moveTo(x_vals[9], y_vals[9], 0.25)
        pyautogui.click()

        # Totals tab
        pyautogui.moveTo(x_vals[10], y_vals[10], 0.25)
        pyautogui.click()

        # Accept button
        # Materials tab
        pyautogui.moveTo(x_vals[11], y_vals[11], 0.25)
        pyautogui.click()

        return

    def execute_wo_releasing_completing_closing(self):
        return

    def exit_program(self):
        self.root.quit()

    def save_wo_nums(self):
        wo_nums = self.entries[3].get("1.0", "end-1c")

        with open("wo_nums.txt", 'w') as file:
            file.write(wo_nums)

        file.close()

        return


if __name__ == "__main__":
    root = tk.Tk()
    app = MyApp(root)
    root.mainloop()
