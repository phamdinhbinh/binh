MF4 Reader v6.1 - project khoi phuc tu file source don
======================================================

Cach nhanh nhat tren Windows:
1. Cai Python 3.11 hoac 3.12, nho chon "Add Python to PATH".
2. Giai nen thu muc nay.
3. Double-click setup_and_run.bat.
4. Lan sau chi can double-click run.bat.

Cau truc bo sung:
  mf4_reader_v6.1.py
  mf4reader/
    __init__.py
    align.py
    cache.py
  requirements.txt
  setup_and_run.bat
  run.bat

Hai module noi bo bi thieu da duoc phuc dung:
- mf4reader.align
  - estimate_offset(): can chinh signal bang correlation, tra ve (offset, quality)
  - trigger_time(): tim rising/falling threshold crossing co noi suy thoi gian
- mf4reader.cache
  - file_hash(): SHA-256 streaming cho file lon

Neu setup_and_run.bat bao loi, chay trong Command Prompt de xem day du log:
  cd <thu_muc_project>
  setup_and_run.bat

Hoac chay thu cong:
  py -m venv .venv
  .venv\Scripts\activate
  python -m pip install -r requirements.txt
  python mf4_reader_v6.1.py
