import shutil
from pathlib import Path

def copy_folders(src_folders, dst_root, ignore_patterns=None):
    for src in src_folders:
        dst = dst_root/src.name
        shutil.copytree(src, dst, ignore=shutil.ignore_patterns(*ignore_patterns))

if __name__ == "__main__":
    SRC1 = Path('test_data')
    SRCs = [SRC1]
    DST = Path('dest_folder')
    copy_folders(SRCs, DST, ignore_patterns=('*.xlsx', '*.docx'))