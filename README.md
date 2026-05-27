# DocX Formatter — Đề án tốt nghiệp

Chuẩn hóa định dạng file .docx theo template đề án tốt nghiệp:
Times New Roman, heading đúng cấp, đánh số hình/bảng theo chương, margin chuẩn.

## Cài đặt & chạy

### Backend (FastAPI)

```bash
cd backend
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```

API chạy tại: http://localhost:8000

### Frontend (React)

Dùng Vite hoặc Create React App:

```bash
cd frontend
npm create vite@latest . -- --template react
# Thay src/App.jsx bằng file App.jsx đã có
npm install
npm run dev
```

Hoặc chạy trực tiếp App.jsx trong Claude artifact (đã tích hợp sẵn).

---

## Endpoints

### POST /format
Upload .docx → trả về .docx đã format

```bash
curl -X POST http://localhost:8000/format \
  -F "file=@de_an.docx" \
  -o de_an_formatted.docx
```

### POST /inspect  
Upload .docx → trả về JSON cấu trúc (để debug)

```bash
curl -X POST http://localhost:8000/inspect \
  -F "file=@de_an.docx"
```

---

## Format Rules (chỉnh tại backend/main.py → STYLE_CONFIG)

| Thành phần      | Quy tắc                                    |
|-----------------|--------------------------------------------|
| Font mặc định   | Times New Roman 14pt                       |
| Line spacing    | 1.5                                        |
| Heading 1       | 14pt, HOA, đậm, căn giữa                  |
| Heading 2       | 14pt, HOA, đậm, căn trái                  |
| Heading 3       | 14pt, đậm, căn trái                        |
| Heading 4       | 14pt, đậm + nghiêng                        |
| Tên hình        | 13pt, nghiêng, căn giữa, dưới hình        |
| Tên bảng        | 13pt, đậm, căn giữa, trên bảng            |
| Đánh số hình    | Hình X.Y. (X = chương, Y = thứ tự)        |
| Đánh số bảng    | Bảng X.Y. (X = chương, Y = thứ tự)        |
| Margin          | T3.5 D3.0 T3.5 P2.0 cm                    |
| Indent đầu dòng | 1.25 cm                                    |
| Font bảng biểu  | Times New Roman 13pt, line spacing 1.2     |

---

## Giới hạn hiện tại

- **Hình vẽ**: giữ nguyên vị trí, KHÔNG di chuyển → phải căn giữa thủ công nếu cần
- **Caption tự động**: chỉ re-number các caption đã có style `Caption`
  - Nếu caption dùng style `Normal` → không nhận dạng được
- **Trang bìa**: giữ nguyên, không reformat (Normal + căn giữa)
- **Header/Footer**: không chỉnh

## Điều chỉnh thêm

Để thêm rule đặc biệt (ví dụ trang bìa, lời cam đoan):
→ Thêm logic vào `format_document()` trong `backend/main.py`

Để nhận dạng caption từ style Normal:
→ Thêm vào `format_document()`:
```python
elif style_name == 'Normal' and RE_FIGURE_CAPTION.match(text):
    para.style = doc.styles['Caption']
    # ... format như caption
```
