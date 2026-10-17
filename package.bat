uv run python build.py
uv run pyinstaller --onedir --windowed --icon=./resources/cropped-favicon-1-270x270.png main.py --name pol-tax --version-file file_version_info.txt
mkdir ".\dist\pol-tax\resources"
copy ".\resources\version.npy" ".\dist\pol-tax\resources\version.npy"
