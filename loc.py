from pywinauto.application import Application
import csv
import time
from datetime import datetime
import pyautogui
import keyboard
import pyperclip
import pytesseract
import cv2
import numpy as np
from PIL import Image, ImageGrab
from tabulate import tabulate  # Import tabulate for formatting
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Global variables
app = None
dlg = None
main_window_x = None
main_window_y = None
fit_window_x = None
fit_window_y = None
circ_window_x = None
circ_window_y = None
over_5_percent_files = []

# Ensure Tesseract OCR is installed and its path is set correctly
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\nikam\PycharmProjects\zviewAuto\tesseract\tesseract.exe'

# Email-to-SMS gateway configuration
SENDER_EMAIL = "bmnldata@gmail.com"  # Replace with your email address
SENDER_PASSWORD = "mhxn pvbx mcql vfxx"  # Replace with your email password
RECIPIENT_EMAIL = str(input("Enter your email: "))


def send_email_alert(csv_file_name):
    message = f"Data set: {csv_file_name} has an error percentage over 5.00%"

    try:
        # Set up the server
        server = smtplib.SMTP("smtp.gmail.com", 587)  # Use Gmail's SMTP server
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)

        # Create the email content
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = RECIPIENT_EMAIL
        msg['Subject'] = "Alert"
        msg.attach(MIMEText(message, 'plain'))

        # Send the message
        server.sendmail(SENDER_EMAIL, RECIPIENT_EMAIL, msg.as_string())
        print(f"Email alert sent to {RECIPIENT_EMAIL}: {message}")

        # Disconnect from the server
        server.quit()
    except Exception as e:
        print(f"Error sending email alert: {e}")

def get_top_left_corner(prompt_message):
    print(prompt_message)
    while True:
        if keyboard.is_pressed('s'):
            x, y = pyautogui.position()
            time.sleep(1)  # Small delay to prevent double capturing
            return x, y

# Example initial known resolution (replace with your tested resolution)
INITIAL_RESOLUTION = (3072, 1920)

# Function to calculate scaling factors
def calculate_scaling_factors(current_resolution):
    scaling_x = current_resolution[0] / INITIAL_RESOLUTION[0]
    scaling_y = current_resolution[1] / INITIAL_RESOLUTION[1]
    return scaling_x, scaling_y

# Function to adjust coordinates using scaling factors
def adjust_coordinates(x, y, scaling_x, scaling_y):
    adjusted_x = int(x * scaling_x)
    adjusted_y = int(y * scaling_y)
    return adjusted_x, adjusted_y

def IF_copy_text(click_coords, top_left_x, top_left_y):
    pyautogui.click(top_left_x + click_coords[0], top_left_y + click_coords[1], button='right')
    pyautogui.click(top_left_x + 815, top_left_y + 22)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(.2)
    return pyperclip.paste()

def paste_text(click_coords, top_left_x, top_left_y, value):
    pyautogui.click(top_left_x + click_coords[0], top_left_y + click_coords[1], button='right')
    pyautogui.click(top_left_x + 1128, top_left_y + 24)
    pyperclip.copy(value)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(.1)

def capture_error_percentage(capture_coords):
    try:
        bbox = (capture_coords[0], capture_coords[1], capture_coords[2], capture_coords[3])
        img = ImageGrab.grab(bbox)

        # Convert PIL image to OpenCV format
        open_cv_image = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)

        # Convert image to grayscale
        gray = cv2.cvtColor(open_cv_image, cv2.COLOR_BGR2GRAY)

        # Apply Gaussian blur to reduce noise
        blurred_img = cv2.GaussianBlur(gray, (5, 5), 0)

        # Increase the contrast of the image
        contrast_img = cv2.addWeighted(blurred_img, 1.5, blurred_img, 0, 0)

        # Apply adaptive thresholding
        threshold_img = cv2.adaptiveThreshold(contrast_img, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)

        # Convert OpenCV image back to PIL format
        pil_image = Image.fromarray(threshold_img)

        # Perform OCR on the processed image
        custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789.-'
        error_percentage = pytesseract.image_to_string(pil_image, config=custom_config)

        # Clean up OCR result
        cleaned_error_percentage = ''.join(filter(lambda x: x.isdigit() or x == '.' or x == '-', error_percentage))

        return cleaned_error_percentage.strip()

    except Exception as e:
        print(f"Error capturing error percentage: {e}")
        return ""

def main():
    global main_window_x, main_window_y, data_window_x, data_window_y, fit_window_x, fit_window_y, circ_window_x, circ_window_y, over_5_percent_files

    print(
        "Please make sure you have loaded the ZView 4 model and data you would like one time before running this script")
    num_files = int(input("Enter the number of data files in the folder: "))
    num_iterations = int(input("Process A: Enter the number of iterations (rounds) to cycle capture error percentages: "))

    # Example usage:
    current_resolution = pyautogui.size()  # Get current screen resolution
    scaling_x, scaling_y = calculate_scaling_factors(current_resolution)

    # Open the app
    app = Application().start("C:\\SAI\\Programs\\ZView4.exe /multicopy")
    time.sleep(3)

    dlg = app.window(title_re="ZView 4")
    dlg.set_focus()

    # Get top-left corner for the main window
    main_window_x, main_window_y = get_top_left_corner("Please mark the top-left corner of the main ZView 4 window with key press 's'")

    initial_x, initial_y = 216, 293
    # Adjust coordinates based on scaling factors
    adjusted_x, adjusted_y = adjust_coordinates(initial_x, initial_y, scaling_x, scaling_y)

    try:
        dlg.type_keys("^d")
        time.sleep(2)

        # Get top-left corner for the data import window
        data_window_x, data_window_y = get_top_left_corner("Please mark the top-left corner of the data import window with key press 's'")

        pyautogui.doubleClick(data_window_x + adjusted_x, data_window_y + adjusted_y, interval=.1)

        for i in range(1, num_files):
            if i <= 15:
                initial_y += 35
            else:
                pyautogui.press('down')

            pyautogui.doubleClick(data_window_x + adjusted_x, data_window_y + adjusted_y, interval=.1)

        print(f"Imported {num_files} data files.")

        # Click OK
        pyautogui.click(data_window_x + 1419, data_window_y + 1278)
        time.sleep(.3)

        # Clicks popup
        pyautogui.press("enter")

        for i in range(num_files):
            # Changes active data
            pyautogui.click(main_window_x + 1861, main_window_y + 172)
            pyautogui.press('down')
            time.sleep(.5)

            # Opens instant fit
            dlg.type_keys("^f")
            time.sleep(.5)

            # Get top-left corner for the fit window if not already set
            if fit_window_x is None or fit_window_y is None:
                fit_window_x, fit_window_y = get_top_left_corner(
                    "Please mark the top-left corner of the instant fit window with key press 's'")

            # Opens equivalent circuits
            dlg.type_keys("^e")
            time.sleep(.2)

            # Opens equivalent circuits after initial run
            if circ_window_x is None or circ_window_y is None:
                circ_window_x, circ_window_y = get_top_left_corner(
                    "Please mark the top-left corner of the equivalent circuits window with key press 's'")

            pyautogui.click(fit_window_x + 234, fit_window_y + 231)
            time.sleep(.1)

            # Copies CPE values
            CPE_coords = [(590, 149), (590, 191), (590, 233)]
            CPE_values = [IF_copy_text(coord, fit_window_x, fit_window_y) for coord in CPE_coords]

            # Open WO instant fit
            pyautogui.click(fit_window_x + 234, fit_window_y + 424)
            time.sleep(.1)

            # Copies WO values
            WO_coords = [(590, 149), (590, 191), (590, 233), (590, 275)]
            WO_values = [IF_copy_text(coord, fit_window_x, fit_window_y) for coord in WO_coords]

            # Paste values in eq circuits
            paste_coords = [(508, 409), (508, 451), (508, 493), (508, 535), (508, 577), (508, 619), (508, 661)]
            paste_values = CPE_values + WO_values

            for i, coord in enumerate(paste_coords):
                paste_text(coord, circ_window_x, circ_window_y, paste_values[i])

            # Close instant fit
            time.sleep(.2)
            dlg.type_keys("^f")

            # Define the coordinates for each button component
            component_click_coords = [
                ('R1', (190, 408)),
                ('CPE1-T', (190, 450)),
                ('CPE1-P', (190, 492)),
                ('R2', (190, 535)),
                ('Wo1-R', (190, 576)),
                ('Wo1-T', (190, 619)),
                ('Wo1-P', (190, 661))
            ]

            # Define the coordinates for capturing the error percentage
            component_capture_coords = [
                (714, 389, 936, 422),  # Adjust based on your setup
                (714, 431, 936, 468),
                (714, 471, 936, 515),
                (714, 517, 936, 550),
                (714, 555, 936, 591),
                (714, 596, 936, 634),
                (714, 638, 936, 680)
            ]

            # Reset the list of files exceeding 5% error at the start of each iteration
            over_5_percent_files.clear()

            has_file_over_5_percent = False

            # Capture error percentage for the final iteration only
            final_percentages = {}
            final_values = []

            for iteration in range(1, num_iterations + 1):
                print(f"\nIteration {iteration}:")

                # Initialize a dictionary to store percentages for this iteration
                iteration_percentages = {}

                for click_coord, capture_coord in zip(component_click_coords, component_capture_coords):
                    component_name, click_position = click_coord
                    pyautogui.click(circ_window_x + click_position[0], circ_window_y + click_position[1])
                    pyautogui.press('r')
                    time.sleep(3)
                    error_percentage = capture_error_percentage(
                        (circ_window_x + capture_coord[0], circ_window_y + capture_coord[1],
                         circ_window_x + capture_coord[2], circ_window_y + capture_coord[3]))

                    if component_name not in iteration_percentages:
                        iteration_percentages[component_name] = []

                    iteration_percentages[component_name].append(error_percentage)
                    time.sleep(1.5)
                    pyautogui.doubleClick(circ_window_x + click_position[0], circ_window_y + click_position[1])

                # Save percentages for the final iteration to the main dictionary
                if iteration == num_iterations:
                    final_percentages = iteration_percentages

                # Print the table for the current iteration
                print(tabulate(iteration_percentages, headers='keys', tablefmt='pretty'))

            # Iterate over each paste coordinate
            for paste_coord in paste_coords:
                pyautogui.click(circ_window_x + paste_coord[0], circ_window_y + paste_coord[1], button='right')
                pyautogui.click(circ_window_x + 1128, circ_window_y + 24)
                pyautogui.hotkey('ctrl', 'c')
                copied_value = pyperclip.paste()

                # Convert copied_value to an integer
                try:
                    float_value = float(copied_value)
                    final_values.append(float_value)
                except ValueError as e:
                    print(f"Error converting '{copied_value}' to an integer: {e}")
            print(final_values)

            # Save final error percentages to CSV file
            timestamp = datetime.now().strftime("%d_%H%M%S")
            csv_file = f"captured_values_{timestamp}.csv"

            for component, percentages in final_percentages.items():
                for percentage in percentages:
                    if float(percentage) > 5.00:
                        if csv_file not in over_5_percent_files:  # Check if file name is not already added
                            over_5_percent_files.append(csv_file)
                        has_file_over_5_percent = True
                        break  # Exit the inner loop once one component is found over 5%

            with open(csv_file, mode='w', newline='') as file:
                writer = csv.writer(file)
                headers = ['Component', 'Error Percentage', 'Value']
                writer.writerow(headers)

                for component, percentages in final_percentages.items():
                    # Get the corresponding final value for the current component
                    final_value = final_values.pop(0)  # Get and remove the first item from final_values
                    row = [component, percentages[0], final_value]  # Assuming percentages is a list
                    writer.writerow(row)
            print(f"Saved results to {csv_file}")

            pyautogui.hotkey('ctrl', 'e')

            # Ensure each file name is added only once
            over_5_percent_files = list(set(over_5_percent_files))
            # Send SMS alert for each dataset exceeding 5% error
            send_email_alert(csv_file)

    except Exception as e:
        print(f"Error: {e}")
    print("Program has completed press 'q' to quit ")

    # Quits app
    while True:
        if keyboard.is_pressed('q'):
            print("Quitting program")
            break
        time.sleep(0.1)
    dlg.close()
#v1
if __name__ == "__main__":
    main()
