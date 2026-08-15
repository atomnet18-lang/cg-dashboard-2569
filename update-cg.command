#!/bin/bash
# =====================================================
#  อัปเดต Dashboard CG 2569 (แปลง Excel -> data.json -> push GitHub)
#  วิธีใช้: ดับเบิลคลิกไฟล์นี้ได้เลย
# =====================================================
cd "$(dirname "$0")" || exit 1

echo "======================================================"
echo "        อัปเดต Dashboard CG ประจำปี 2569"
echo "======================================================"
echo ""

XLSX="CG_2569_บันทึกผลจริง_template.xlsx"

if [ ! -f "$XLSX" ]; then
  echo "  [X] ไม่พบไฟล์ Excel: $XLSX"
  echo "      กรุณาวางไฟล์ไว้ในโฟลเดอร์เดียวกับสคริปต์นี้"
  echo ""
  read -p "กด Enter เพื่อปิดหน้าต่าง..."
  exit 1
fi

echo "  [1/3] แปลง Excel เป็น data.json ..."
if ! python3 convert_excel_to_json.py "$XLSX" data.json; then
  echo "  [X] แปลงไม่สำเร็จ (ตรวจว่าติดตั้ง openpyxl แล้ว: python3 -m pip install --user --break-system-packages openpyxl)"
  echo ""
  read -p "กด Enter เพื่อปิดหน้าต่าง..."
  exit 1
fi
echo ""

echo "  [2/3] ตรวจการเปลี่ยนแปลง ..."
git add -A
if git diff --cached --quiet; then
  echo "  [i] ไม่มีข้อมูลเปลี่ยนแปลง ไม่ต้องอัปเดต"
  echo ""
  read -p "กด Enter เพื่อปิดหน้าต่าง..."
  exit 0
fi
git --no-pager diff --cached --stat
echo ""

echo "  [3/3] ส่งขึ้น GitHub ..."
git commit -m "อัปเดตข้อมูล CG $(date '+%Y-%m-%d %H:%M')" >/dev/null
if git push; then
  echo ""
  echo "  [OK] อัปเดตสำเร็จ! เว็บ GitHub Pages จะเปลี่ยนภายใน 1-2 นาที"
else
  echo ""
  echo "  [X] push ไม่สำเร็จ - ตรวจอินเทอร์เน็ต หรือเปิด GitHub Desktop กด Push แทน"
fi

echo ""
echo "======================================================"
read -p "เสร็จแล้ว กด Enter เพื่อปิดหน้าต่าง..."
