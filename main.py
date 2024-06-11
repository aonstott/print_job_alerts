import sys
from Emailer import Emailer


def run(filename, send_emails):
    try:
        emailer = Emailer(filename)
        emailer.assemble_excel_files(send_emails)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) != 3 and len(sys.argv) != 2:
        print("Usage: python your_script.py filename send_emails")
        sys.exit(1)
    
    filename = sys.argv[1]
    if len(sys.argv) == 3:
        send_emails = sys.argv[2]
    run(filename, send_emails)


