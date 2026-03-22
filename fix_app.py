import sys

def fix_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Replace literal \` with `
    content = content.replace('\\`', '`')
    # Replace literal \$ with $
    content = content.replace('\\$', '$')
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)

if __name__ == "__main__":
    fix_file(sys.argv[1])
