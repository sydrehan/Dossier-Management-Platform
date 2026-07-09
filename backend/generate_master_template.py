import sys
from pathlib import Path

# Import our new templates library
sys.path.append(str(Path(__file__).parent.resolve()))
import generate_templates

def generate_master_template():
    generate_templates.build_templates()

if __name__ == "__main__":
    generate_master_template()
