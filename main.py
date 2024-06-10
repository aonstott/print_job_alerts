import sys
from Emailer import Emailer


def run(filename):
    try:
        emailer = Emailer(filename)
        emailer.assemble_excel_files()
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python your_script.py filename")
        sys.exit(1)
    
    filename = sys.argv[1]
    run(filename)


