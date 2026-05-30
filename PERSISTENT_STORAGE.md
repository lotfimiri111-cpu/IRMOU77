# إصلاح التخزين المؤقت على Render Free

## المشكلة
Render Free يمسح `/opt/render/project/src/` عند كل نشر جديد.
كل طلب (order) وكل وصل (receipt) وكل ملف PPTX كان يضيع.

## الحل المطبّق (v29)

### 1. Render Persistent Disk
في `render.yaml` أُضيف:
```yaml
disk:
  name: mathkarati-data
  mountPath: /data
  sizeGB: 1
```
كل البيانات تُحفظ على `/data` الذي **لا يُمسح** عند النشر.

### 2. متغيرات البيئة المحدّثة
```
STORAGE_DIR = /data/storage
DB_PATH     = /data/mathkarati_payments.db
```

### 3. Migration تلقائي
`app.py` يفحص عند الإقلاع:
- إذا وُجد DB قديم في مجلد المشروع → ينسخه إلى `/data`
- إذا وُجد storage قديم → ينسخه إلى `/data/storage`

### 4. PPTX محمية بعد الاعتماد
`store_pptx()` الآن تنسخ الملف إلى:
`/data/storage/pptx/orders/<order_id>.pptx`
هذا المسار دائم وبقى حتى لو جُدِّد التطبيق.

---

## خطوات التفعيل على Render

1. **ادفع الكود** → Render يقرأ `render.yaml` الجديد
2. في Dashboard: **Settings → Disks** → ستظهر `mathkarati-data` تلقائياً
3. تحقق من `/health` → يجب أن يرجع `"persistent_disk": true`

## ماذا لو كان عندك بيانات قديمة؟

البيانات القديمة **ضاعت مع كل نشر سابق** — هذا ما كانت المشكلة.
لكن من الآن فصاعداً، أي طلب جديد محمي بالكامل.

## تحقق سريع بعد النشر
```
GET /health
→ { "persistent_disk": true, "db_path": "/data/mathkarati_payments.db", ... }
```

## ملاحظة عن Render Free Disk
- Render يقدم 1GB disk مجاني مع Free plan
- الـ disk لا يُمسح عند النشر، لكن قد يُمسح إذا حذفت الـ service نهائياً
- لذا: أخذ نسخ احتياطية دورية من `/data/mathkarati_payments.db` مستحسن
