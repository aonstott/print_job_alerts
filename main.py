import sys

def process_file(filename):
    try:
        with open(filename, 'r') as file:
            contents = file.read()
            print("File contents:")
            print(contents)
    except FileNotFoundError:
        print("File not found:", filename)
    except Exception as e:
        print("Error:", e)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python your_script.py filename")
        sys.exit(1)
    
    filename = sys.argv[1]
    process_file(filename)


