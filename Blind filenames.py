from os import listdir, getcwd, rename
from os.path import isfile, join
from natsort import os_sorted  # Import for sorting files naturally (unused here)
from random import shuffle
from datetime import date

# Set up variables
today = date.today()  # Get today's date
ext = '.czi'  # File extension to filter by

cwd = getcwd()  # Get the current working directory

# Get a list of files in the current directory that have the specified extension
to_rename = [f for f in listdir(cwd) if (isfile(join(cwd, f)) and f.endswith(ext))]

# Create a list of randomized numbers (as strings) equal to the number of files to rename
rand = [str(i + 1) for i in range(len(to_rename))]
shuffle(rand)  # Shuffle the numbers to randomize the order

# Initialize a key list to store the renaming key information
key = [f'Randomized key originated at:\nDir: {cwd}\nDate: {today.strftime("%d/%m/%Y")}\n\n']

# Renaming process
if __name__ == "__main__":
    for i in range(len(to_rename)):
        # Append the original file name and its new name to the key list
        key.append(f'{rand[i]} = {to_rename[i].strip(ext)}\n')

        # Rename the file with the new randomized name
        rename(to_rename[i], rand[i] + ext)

    # Write the key information to a text file for reference
    with open("key.txt", "w") as f:
        for ele in key:
            f.write(ele)
