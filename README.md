# ⚖️ AB Layout (Advocate Benefit Layout)

![Version](https://img.shields.io/badge/Version-1.0-blue)
![License](https://img.shields.io/badge/License-Apache%202.0-green)
![Platform](https://img.shields.io/badge/Platform-Microsoft%20Word%20Add--in-orange)
![Status](https://img.shields.io/badge/Status-Stable-brightgreen)

---

## 📌 About

**AB Layout (Advocate Benefit Layout)** is an open-source **Microsoft Word Office JS Add-in** designed for **Indian legal drafting**.

It enables advocates, law students, and professionals to:

* ⚡ Apply **court-specific formatting instantly**
* 📄 Convert documents into **Legal (8.5" × 14") layouts**
* 🔁 Toggle layouts with **Undo support**
* 🌐 Stay updated via **GitHub-powered layout sync**

---

## 🚀 Features

* 🏛 **Multi-Court Layout Support** (High Court, District Court, etc.)
* ⚙️ **One-Click Formatting**
* 🔁 **Undo Layout Changes**
* 📑 Works on **existing documents (even 100+ pages)**
* 🌐 **Live Updates via GitHub (no reinstall needed)**
* 🧩 **Open Source & Extensible**

---

## 🖥️ Demo / Live

🌐 **Live Add-in Source:**
👉 https://isrgrajan.github.io/ab-layout/

---

## 📥 Installation Guide

### Step 1: Download Manifest

Download the `manifest.xml` file from the repository.

### Step 2: Open Microsoft Word

Go to:

```
Insert → Add-ins → Upload My Add-in
```

### Step 3: Upload Manifest

Select the downloaded `manifest.xml`

🎉 Done! AB Layout will appear inside Word.

---

## 🧭 How to Use

1. Open the **AB Layout panel**
2. Select:

   * State
   * Court
3. Click:

   * ✅ **Apply Layout**
   * 🔁 **Undo** (if needed)

---

## 📁 Project Structure

```bash
ab-layout/
│
├── manifest.xml
├── web/
│   ├── taskpane.html
│   ├── taskpane.js
│   ├── styles.css
│
├── layouts/
│   └── layouts.json
│
├── assets/
│   └── icon-32.png
│
└── README.md
```

---

## 🔗 Internal Navigation

* 📌 [About](#-about)
* 🚀 [Features](#-features)
* 📥 [Installation Guide](#-installation-guide)
* 🧭 [How to Use](#-how-to-use)
* 📁 [Project Structure](#-project-structure)
* 🤝 [Contributing](#-contributing)
* 📄 [License](#-license)

---

## 🤝 Contributing

We welcome contributions from the legal and developer community!

### ➕ Add a New Court Layout

1. Go to `/layouts/layouts.json`
2. Add your court format
3. Submit a Pull Request

### 📌 Ensure:

* Correct margins
* Proper naming
* Tested formatting in Word

---

## 📄 License

Licensed under the **Apache License 2.0**

© 2026 **Bee Isrg Rajan**

---

## 🌟 Support the Project

If you find this useful:

* ⭐ Star the repository
* 🍴 Fork and contribute
* 📢 Share with legal professionals

---

## 📬 Contact / Maintainer

**Maintained by:** Bee Isrg Rajan
🔗 GitHub: https://github.com/isrgrajan

---

## 🔍 SEO Keywords

> Microsoft Word Add-in, Legal Document Formatting, Indian Court Format, Advocate Tools, High Court Formatting, Legal Drafting Automation, Office JS Add-in, Court Layout Tool

---
