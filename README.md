# generate_ppt

## 1. Set up
1. 이전 실습 폴더 (`BEFORE_DIR`) 생성 (ex. 0305)
2. 이번 실습 폴더 생성 (`TODAY_DIR`) (ex. 0312)
3. `BEFORE_DIR` - `PPT` 에 이전 PPT 저장
4. `TODAY_DIR` - `사진` 에 이번 사진 저장
5. `generate_ppt.py` 에 폴더명 명시
```python
BEFORE_DIR = "0305"
TODAY_DIR = "0312"
```

## 2. Run
``` zsh
source path/to/venv/bin/activate
python3 generate_ppt.py

(실행이 안되면)
pip3 install pptx
python3 generate_ppt.py
```
