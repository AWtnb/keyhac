import os
import shutil
import subprocess
from pathlib import Path


def make_junction() -> None:
    d = ".keyhac"
    data_path = Path(os.environ["USERPROFILE"]) / d
    src_path = Path(__file__).resolve().parent / d

    if data_path.is_symlink() or os.path.isjunction(data_path):
        os.rmdir(data_path)
    elif data_path.is_dir():
        shutil.rmtree(data_path, ignore_errors=True)
    elif data_path.exists():
        data_path.unlink()

    # ジャンクション作成
    subprocess.run(
        ["cmd", "/c", "mklink", "/J", str(data_path), str(src_path)],
        check=True,
    )


if __name__ == "__main__":
    make_junction()
