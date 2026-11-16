# lamkq-marketplace — plugin Claude Code cho dự án "Làm KQ"

Plugin chỉ đóng gói **hành vi của Claude** (skill, hook). Code dự án (`Core/`, template, `HieuChuan/`, `CongCu/`) vẫn phải có sẵn trên từng PC — chép cả thư mục `Làm KQ` (hoặc clone kho git khi xong việc V9).

## Có gì

| Thành phần | Việc |
|---|---|
| Hook `SessionStart` (`scripts/dau_phien.py`) | Khi mở phiên trong dự án Làm KQ: đặt `UPCHUNGTHU_PATH`, `DRAWIO_TOOL_DIR` trỏ vào `CongCu\` và chạy `python Core/Tools/context.py dau-phien`, đưa kết quả vào ngữ cảnh. Dự án khác: không làm gì. |
| Hook `PreToolUse` (`scripts/chan_lenh_z.py`) | Chặn `robocopy /MIR`, `/PURGE` và lệnh xóa / di chuyển / đổi tên trên ổ Z (AGENTS.md mục 0). |
| Skill `lam-ho-so` | Thứ tự làm một hồ sơ: khóa → đọc → dựng → hỏi gộp → QA → chép Z → ghi nhật ký → trả khóa. Kèm ghi chú khách quen (Nextern). |
| Skill `chep-z` | Chép ThucHien / TBKQ / pdf lên Z bằng `ai_files.py`, xử lý `[Z EDITED]`, `[Z Duplicate]`. |
| Skill `setup-may-moi` | Kiểm tra và cài môi trường trên PC mới (thay cho dán prompt `SETUP-MOI-TRUONG-MOI.md`). |

`AGENTS.md` ở gốc dự án vẫn là nguồn quy tắc duy nhất — skill chỉ trỏ về các mục của nó.

## Cài trên mỗi PC

Cần: Claude Code, Python trên PATH (hoặc biến `LAMKQ_PYTHON`), thư mục dự án `Làm KQ` đã có `CongCu\lamkq-marketplace`.

Chạy trong terminal (dialog `/plugin` không có trong app desktop). Cách 1 — từ GitHub (repo Private, máy đã đăng nhập GitHub; dùng `.claude-plugin/marketplace.json` ở gốc repo):

```
claude plugin marketplace add phatnguyen1999p/Make-Document-QUATEST3
claude plugin install lamkq@lamkq-marketplace
```

Cách 2 — từ thư mục dự án đã có trên máy (máy N2-PHAT cài theo cách này 16/11/2026):

```
claude plugin marketplace add "C:\Users\<user>\Desktop\Source\Làm KQ\CongCu\lamkq-marketplace"
claude plugin install lamkq@lamkq-marketplace
```

Hai cách cùng tên marketplace `lamkq-marketplace` → mỗi máy chỉ dùng một cách.

Mở phiên mới trong thư mục dự án → thấy dòng `[lamkq] Công cụ dùng bản trong CongCu:` là hook đã chạy.

Cập nhật sau khi sửa plugin (tăng `version` trong cả `plugin.json` và `marketplace.json`):

```
claude plugin marketplace update lamkq-marketplace
claude plugin update lamkq@lamkq-marketplace
```

### Cách khác: tự nhắc cài khi mở dự án

Thêm vào `.claude/settings.json` ở gốc dự án (file này đi theo dự án sang mọi máy):

```json
{
  "extraKnownMarketplaces": {
    "lamkq-marketplace": {
      "source": { "source": "directory", "path": "./CongCu/lamkq-marketplace" }
    }
  },
  "enabledPlugins": { "lamkq@lamkq-marketplace": true }
}
```

## Giới hạn khi nhiều PC cùng làm (việc V9)

- Khóa hồ sơ `Context/_khoa/` và nhật ký `Context/` nằm trên từng máy → máy khác không thấy khóa.
- Sổ `ai_files.json` nằm trong `Input\<6 số>\ThucHien\_checkpoint\` của máy đã dựng → máy khác thấy file trên Z là `[Z EDITED]`.

Cho đến khi sửa xong: **mỗi hồ sơ chỉ làm trên một máy**.
