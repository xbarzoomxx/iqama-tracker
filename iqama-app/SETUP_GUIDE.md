# 🚀 دليل نشر نظام متابعة الإقامات

## الخطوات (تستغرق ~15 دقيقة)

---

## ① إنشاء مشروع Firebase (مجاني)

1. افتح **https://console.firebase.google.com**
2. اضغط **"Create a project"** → اكتب اسم مثل `iqama-anjal`
3. أوقف Google Analytics (اختياري) → **Create project**
4. من القائمة الجانبية اختر **Build → Realtime Database**
5. اضغط **"Create Database"** → اختر منطقة **us-central1** → **Start in test mode** → **Enable**
6. من القائمة الجانبية اختر **Project Settings** (⚙️ بجانب Project Overview)
7. انزل لـ **"Your apps"** → اضغط أيقونة **`</>`** (Web)
8. اكتب اسم التطبيق → **Register app**
9. **انسخ** كامل كود `firebaseConfig` الذي يظهر

---

## ② وضع بيانات Firebase في التطبيق

افتح ملف `src/App.jsx` وابحث عن:

```js
const firebaseConfig = {
  apiKey: "REPLACE_API_KEY",
  authDomain: "REPLACE_AUTH_DOMAIN",
  databaseURL: "REPLACE_DATABASE_URL",
  ...
};
```

**استبدل** كل قيمة `REPLACE_...` بالقيم التي نسختها من Firebase.

> ⚠️ **مهم:** `databaseURL` تبدو هكذا:
> `https://iqama-anjal-default-rtdb.firebaseio.com`
> إذا لم تظهر، اذهب لـ Realtime Database وانسخها من أعلى الصفحة.

---

## ③ رفع المشروع على GitHub

1. افتح **https://github.com** → سجّل دخول أو أنشئ حساباً
2. اضغط **"New repository"** → اسم مثل `iqama-tracker` → **Create**
3. حمّل مجلد المشروع كاملاً (drag & drop أو Git):

```bash
# في terminal داخل مجلد المشروع:
git init
git add .
git commit -m "initial commit"
git branch -M main
git remote add origin https://github.com/USERNAME/iqama-tracker.git
git push -u origin main
```

---

## ④ النشر على Vercel (مجاني)

1. افتح **https://vercel.com** → سجّل دخول بحساب GitHub
2. اضغط **"Add New Project"**
3. اختر Repository `iqama-tracker` → **Import**
4. Framework: **Vite** (يكتشفه تلقائياً)
5. اضغط **Deploy** ✅
6. بعد دقيقة يعطيك رابطاً مثل: `https://iqama-tracker-xyz.vercel.app`

---

## ⑤ تثبيته على الجوال (PWA)

### iOS (iPhone):
1. افتح الرابط في **Safari**
2. اضغط أيقونة المشاركة 📤
3. اختر **"Add to Home Screen"**
4. اضغط **Add** ✅

### Android:
1. افتح الرابط في **Chrome**
2. ستظهر رسالة **"Add to Home Screen"** تلقائياً
3. أو اضغط ⋮ → **Install app**

---

## ⑥ أيقونات التطبيق (اختياري)

ضع في مجلد `public/icons/`:
- `icon-192.png` (192×192 بكسل)
- `icon-512.png` (512×512 بكسل)

يمكنك استخدام شعار شركة أنجال المشاعر.

---

## ✅ النتيجة

| الميزة | التفاصيل |
|--------|----------|
| 🌐 رابط ثابت | `https://your-app.vercel.app` |
| 📱 PWA على الجوال | أيقونة على الشاشة الرئيسية |
| 🔥 تزامن فوري | أي تعديل يظهر فوراً على كل الأجهزة |
| 💰 التكلفة | **مجاني بالكامل** |

---

## 🔒 تأمين البيانات (موصى به)

في Firebase Console → Realtime Database → Rules، غيّر القواعد لـ:

```json
{
  "rules": {
    ".read": "auth != null",
    ".write": "auth != null"
  }
}
```

ثم فعّل **Authentication** → Email/Password وأنشئ حساباً للمستخدمين المصرّح لهم.

---

## 🆘 مشاكل شائعة

**"databaseURL is missing"** → أضف `databaseURL` يدوياً من Realtime Database.

**البيانات لا تتزامن** → تأكد أن rules في Firebase على `test mode`.

**الأيقونة لا تظهر على iOS** → استخدم Safari وليس Chrome لإضافتها.
