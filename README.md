# AutoDistribute

AutoDistribute adalah macro VBA untuk CorelDRAW yang dirancang untuk mengatur dan mendistribusikan objek **Design** dan **Cut Line** secara otomatis ke page dan layer yang sesuai.

Macro ini dibuat untuk mengurangi pekerjaan manual saat menyiapkan file produksi cutting, terutama ketika banyak objek harus dipindahkan, dipasangkan, diposisikan, atau disusun ke beberapa page dengan aturan yang konsisten.

AutoDistribute mendukung dua workflow utama, yaitu **Kiss A** dan **Die A**, serta menyediakan pengaturan tambahan seperti **Sequentially**, rotasi objek, page size, dan pengelolaan layer sensor.

## Main Features

### Object Container

Objek yang akan diproses terlebih dahulu dimasukkan ke dalam antrean sebagai:

* `Design`
* `Cut Line`

Setiap objek disimpan berdasarkan `StaticID`, sehingga macro dapat mencari kembali objek yang sama ketika proses dijalankan.

Antrean ditampilkan melalui `lbxObjects` dengan format seperti:

`Design | Obj. ID: /12345`

atau:

`Cut Line | Obj. ID: /12346`

Objek dalam group maupun Master Page tetap dapat ditemukan selama `StaticID` masih tersedia.

Antrean hanya dapat berisi objek dari satu dokumen dan akan dikosongkan ketika UserForm ditutup.

---

## Processing Modes

### Kiss A

Mode **Kiss A** digunakan ketika Design dan Cut Line perlu ditempatkan sebagai pasangan pada page yang sama.

Secara umum:

* Design ditempatkan pada `Layer 1`.
* Cut Line ditempatkan pada `Layer 2`.
* Design dan Cut Line diposisikan ke tengah page.
* Cut Line diberi outline **CMYK 0, 100, 100, 0**.
* Cut Line tetap aktif untuk print/export.
* Design pada page yang memiliki Cut Line dinonaktifkan dari print/export.

Tanpa `Sequentially`, jumlah Design dan Cut Line harus sama.

Contoh antrean:

`Design 1, Design 2, Cut Line 1, Cut Line 2`

akan diproses sebagai:

`Page 1 → Design 1 + Cut Line 1`

`Page 2 → Design 2 + Cut Line 2`

Urutan kategori di dalam antrean tidak digunakan sebagai pasangan langsung. Macro mengumpulkan Design dan Cut Line, kemudian memasangkannya berdasarkan urutan masing-masing.

### Kiss A + Sequentially

Ketika `Sequentially` aktif, urutan antrean menjadi bagian dari aturan pemrosesan.

Contoh:

`D1, D2, C1, D3, D4, D5, C2`

akan menghasilkan:

`Page 1 → D1 + C1`

`Page 2 → D2`

`Page 3 → D3 + C2`

`Page 4 → D4`

`Page 5 → D5`

Cut Line dipasangkan dengan Design pertama sejak awal antrean atau sejak Cut Line sebelumnya.

Design yang tidak memiliki pasangan Cut Line tetap diproses dan tetap aktif untuk print/export.

Cut Line yang tidak memiliki Design sebelumnya pada segment yang sama akan dianggap sebagai antrean yang tidak valid.

---

### Die A

Mode **Die A** menempatkan setiap objek pada page terpisah.

Tanpa `Sequentially`, Design dan Cut Line dipasangkan berdasarkan urutan masing-masing, tetapi setiap objek tetap mendapat page sendiri.

Contoh:

`Design 1, Design 2, Cut Line 1, Cut Line 2`

diproses menjadi:

`Page 1 → Design 1`

`Page 2 → Cut Line 1`

`Page 3 → Design 2`

`Page 4 → Cut Line 2`

Cut Line pada mode Die A diberi outline **RGB 255, 0, 255** (`#FF00FF`).

Objek ditempatkan pada `Layer 1` dan tetap aktif untuk print/export.

Setelah proses selesai, macro memeriksa layer lokal bernama `Layer 2`, `Layer 3`, dan seterusnya. Layer tersebut hanya akan dihapus apabila kosong. Layer yang masih memiliki objek akan dipertahankan dan dilaporkan kepada user.

### Die A + Sequentially

Ketika `Sequentially` digunakan pada Die A, macro mengikuti urutan `lbxObjects` secara langsung.

Setiap entry memperoleh satu page sendiri sesuai posisi entry tersebut di dalam antrean.

Mode ini berguna ketika urutan Design dan Cut Line sudah ditentukan langsung oleh user dan tidak perlu dipasangkan kembali oleh macro.

---

## Rotation

AutoDistribute menyediakan dua pilihan rotasi:

* **CW -90°** — memutar objek 90° searah jarum jam.
* **CCW +90°** — memutar objek 90° berlawanan arah jarum jam.

Kedua pilihan bersifat saling eksklusif.

Jika tidak ada pilihan yang aktif, objek tidak akan diputar.

Setelah rotasi dilakukan, objek akan diposisikan ke tengah page tujuan.

---

## Page Setup

AutoDistribute menyediakan beberapa pengaturan page yang dapat digunakan langsung dari Main UserForm.

### Default Size

Mengatur Master Page dan seluruh page dokumen menjadi:

`325 × 485 mm`

### Extended Size

Mengatur Master Page dan seluruh page dokumen menjadi:

`335 × 487 mm`

Jika proses distribusi membutuhkan page tambahan, AutoDistribute akan membuat page baru dengan ukuran:

`335 × 487 mm`

---

## MasterPage

`MasterPage` mengatur page menjadi `325 × 485 mm` dan memindahkan layer sensor lokal dari Active Page menjadi Master Layer.

Layer sensor yang diproses adalah:

* `scpro2_printmargin`
* `scpro2_regmarks`
* `scpro2_printonly`

Urutan layer dijaga dari atas ke bawah sebagai:

`scpro2_printmargin`

`scpro2_regmarks`

`scpro2_printonly`

Setelah page setup selesai, `Layer 1` pada page lokal diaktifkan untuk print/export, sedangkan `Layer 2` dinonaktifkan.

---

## PerPage

`PerPage` mengembalikan layer sensor dari status Master menjadi layer lokal pada Active Page.

Ukuran page juga diatur menjadi:

`325 × 485 mm`

Urutan ketiga layer sensor tetap dipertahankan dengan aturan yang sama seperti pada `MasterPage`.

Setelah proses selesai, `Layer 1` pada page lokal diaktifkan untuk print/export dan `Layer 2` dinonaktifkan.

---

## Basic Workflow

1. Buka dokumen CorelDRAW yang ingin diproses.
2. Jalankan AutoDistribute.
3. Seleksi satu atau beberapa objek Design.
4. Tekan **Set Design**.
5. Seleksi satu atau beberapa objek Cut Line.
6. Tekan **Set Cut Line**.
7. Periksa urutan objek pada `lbxObjects`.
8. Pilih **Kiss A** atau **Die A**.
9. Aktifkan `Sequentially` jika urutan antrean ingin digunakan sebagai aturan pemrosesan.
10. Pilih rotasi jika diperlukan.
11. Gunakan pengaturan page atau sensor layer jika diperlukan.
12. Tekan **Process**.

AutoDistribute akan memvalidasi antrean dan dokumen sebelum memindahkan objek.

---

## Object Queue

AutoDistribute menggunakan `StaticID` untuk menyimpan referensi objek, bukan teks yang tampil pada ListBox.

Karena itu, `lbxObjects` hanya berfungsi sebagai representasi visual dari antrean.

Beberapa aturan penting:

* Objek yang sama tidak ditambahkan dua kali ke antrean.
* Antrean hanya dapat menggunakan satu dokumen sumber.
* Dokumen sumber harus aktif ketika `Process` dijalankan.
* Objek yang tidak lagi dapat ditemukan melalui `StaticID` akan membatalkan proses.
* Pencarian mencakup page biasa, Master Page, dan objek di dalam group.
* **Remove** menghapus entry terpilih dari antrean tanpa menghapus objek CorelDRAW.
* **Clear** mengosongkan seluruh antrean tanpa menghapus objek CorelDRAW.
* Menutup UserForm juga mengosongkan antrean.

---

## Processing Safety

Sebelum memodifikasi dokumen, AutoDistribute terlebih dahulu mencoba menyelesaikan referensi seluruh objek dan membangun execution plan.

Perubahan dokumen dijalankan di dalam CorelDRAW `CommandGroup` agar operasi lebih terkontrol dan dapat ditangani sebagai satu kelompok perubahan.

Macro juga berusaha memulihkan:

* Active Page awal.
* Unit dokumen.
* Status `CommandGroup` ketika terjadi error.

Untuk operasi page setup, macro juga memeriksa keberadaan layer tujuan dan jumlah objek setelah perpindahan layer sensor.

Jika kondisi yang tidak aman ditemukan, proses dihentikan dan informasi operasi yang gagal ditampilkan kepada user.

---

## Project Structure

### `AutoDistributeMenu.vba`

Code-behind untuk Main UserForm.

Mengatur interaksi UI seperti:

* pemilihan Kiss A / Die A;
* CW / CCW rotation;
* Sequentially;
* Set Design / Set Cut Line;
* Remove / Clear;
* Default Size / Extended Size;
* MasterPage / PerPage;
* Process.

### `AutoDistributeObjectStore.bas`

Core engine AutoDistribute.

Modul ini menangani:

* penyimpanan Object Container;
* pencarian objek berdasarkan `StaticID`;
* execution planning;
* distribusi page dan layer;
* Kiss A / Die A;
* Sequentially;
* rotation dan centering;
* Cut Line outline;
* page setup;
* sensor layer conversion;
* cleanup layer;
* validation dan error handling.

### `AutoDistributeUF.bas`

Entry point untuk membuka AutoDistribute UserForm secara `vbModeless`.

Entry point:

`AutoDistributeUFMenu`

### `Changelog.log`

Mencatat penambahan fitur, perubahan behavior, perbaikan, dan perkembangan AutoDistribute.

---

## Notes

AutoDistribute dibuat untuk workflow produksi tertentu di CorelDRAW. Beberapa behavior—terutama nama layer, ukuran page, warna Cut Line, dan aturan Kiss A / Die A—bersifat spesifik terhadap workflow tersebut.

Macro tidak dimaksudkan sebagai generic page distribution engine untuk seluruh kemungkinan struktur dokumen CorelDRAW.

Perubahan pada nama layer atau struktur dokumen dapat memerlukan penyesuaian pada source code.

---

## Current Scope

AutoDistribute saat ini mencakup:

* Object Container berbasis `StaticID`;
* Design dan Cut Line roles;
* Kiss A;
* Die A;
* Sequential processing;
* CW / CCW rotation;
* automatic centering;
* automatic page distribution;
* automatic page creation;
* Default dan Extended page size;
* MasterPage / PerPage sensor management;
* Master Page dan nested group object lookup;
* Cut Line color assignment;
* layer print/export handling;
* empty Die layer cleanup;
* validation dan operation-specific error reporting.

Pengembangan berikutnya dapat memperluas workflow ini tanpa mengubah fungsi utama AutoDistribute sebagai alat untuk mengurangi pekerjaan distribusi objek secara manual di CorelDRAW.

Source boleh dipelajari dan dikembangkan, dan issue/feedback tentang bug, edge case, CorelDRAW API, architecture, atau improvement sangat dihargai.
