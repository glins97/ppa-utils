import subprocess
from scripts.generation import generate_csv
import os

if __name__ == "__main__":
    csv = generate_csv()
    for fn in os.listdir("inputs/"):
        if fn.endswith(".docx") and 'RFS' not in fn:
            subprocess.call(['libreoffice', '--headless', '--convert-to',  'pdf', os.path.join("inputs/", fn), '--outdir', 'outputs'])

