import sys
import os
FILE_DIR = os.path.abspath('.')
sys.path.append(FILE_DIR)


from processor.guis.process_gui import run_gui, root

def run(root):
   run_gui(root)

if (__name__ == '__main__'):
   print(sys.path)
   run(root)