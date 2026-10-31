<?xml version="1.0" encoding="UTF-8" ?>
<!DOCTYPE html>
<html b:css='false' b:defaultwidgetversion='2' b:layoutsVersion='3' b:responsive='true' xmlns='http://www.w3.org/1999/xhtml' xmlns:b='http://www.google.com/2005/gml/b' xmlns:data='http://www.google.com/2005/gml/data' xmlns:expr='http://www.google.com/2005/gml/expr'>
<head>
    <meta charset="UTF-8"/>
    <meta content="width=device-width, initial-scale=1.0" name="viewport"/>
    <title><data:blog.pageTitle/></title>
    
    <!-- Dependencies -->
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&amp;display=swap" rel="stylesheet"/>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css"/>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    
    <script>
    //<![CDATA[
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: { sans: ['Inter', 'sans-serif'] },
                    colors: { primary: '#4f46e5', secondary: '#0f172a' }
                }
            }
        }
    //]]>
    </script>

    <b:skin><![CDATA[
        #app-container {
            position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
            background-color: #f8fafc; z-index: 9999999;
            display: flex; flex-direction: column; overflow: hidden;
        }
        .scrollable { overflow-y: auto; height: 100%; }
        .hidden { display: none !important; }
        .modal-overlay { background-color: rgba(15, 23, 42, 0.6); backdrop-filter: blur(4px); }
        
        /* Custom Scrollbar */
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: #f1f1f1; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 10px; }
        ::-webkit-scrollbar-thumb:hover { background: #94a3b8; }

        /* Animation */
        .fade-in { animation: fadeIn 0.3s ease-in-out; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        
        /* Menyembunyikan elemen bawaan Blogger */
        #blogger-mandatory-widgets { display: none !important; }
    ]]></b:skin>
</head>
<body class="text-slate-800 antialiased">

    <div id="app-container">
        
        <!-- Input File Hidden untuk Import Excel -->
        <input type="file" id="excel-file" accept=".xlsx, .xls" class="hidden" onchange="handleImportExcel(event)"/>

        <!-- LOGIN SCREEN -->
        <div id="login-screen" class="flex items-center justify-center h-full bg-slate-900 relative overflow-hidden">
            <!-- Background Decoration -->
            <div class="absolute w-96 h-96 bg-primary rounded-full blur-3xl opacity-20 -top-20 -left-20"></div>
            <div class="absolute w-96 h-96 bg-indigo-400 rounded-full blur-3xl opacity-20 bottom-0 right-0"></div>
            
            <div class="bg-white p-10 rounded-2xl shadow-2xl w-full max-w-md z-10 fade-in border border-slate-100">
                <div class="text-center mb-8 w-full flex flex-col items-center">
                    <div id="login-logo-container" class="w-full flex justify-center"></div>
                    <h2 class="text-3xl font-bold text-slate-900 text-center w-full" id="login-school-name">EduAdmin</h2>
                    <p class="text-slate-500 mt-1 text-center w-full">Sistem Informasi Akademik</p>
                </div>
                <div id="login-alert" class="hidden mb-6 p-3 bg-red-50 border-l-4 border-red-500 text-red-700 rounded text-sm text-center font-medium"></div>
                <form id="login-form" onsubmit="handleLogin(event)">
                    <div class="mb-5">
                        <label class="block text-slate-700 text-sm font-semibold mb-2">Username</label>
                        <div class="relative">
                            <i class="fa-solid fa-user absolute left-4 top-3.5 text-slate-400"></i>
                            <input type="text" id="username" class="w-full pl-11 pr-4 py-3 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-primary focus:border-transparent transition bg-slate-50" placeholder="Masukkan Username" required="required"/>
                        </div>
                    </div>
                    <div class="mb-8">
                        <label class="block text-slate-700 text-sm font-semibold mb-2">Password</label>
                        <div class="relative">
                            <i class="fa-solid fa-lock absolute left-4 top-3.5 text-slate-400"></i>
                            <input type="password" id="password" class="w-full pl-11 pr-4 py-3 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-primary focus:border-transparent transition bg-slate-50" placeholder="Masukkan Password" required="required"/>
                        </div>
                    </div>
                    <button type="submit" class="w-full bg-primary text-white font-semibold py-3 px-4 rounded-xl hover:bg-indigo-700 focus:ring-4 focus:ring-indigo-200 transition duration-300 shadow-lg shadow-indigo-200" id="btn-login">
                        Login ke Portal
                    </button>
                    <div id="demo-info" class="mt-6 p-4 bg-slate-50 rounded-xl border border-slate-100 text-xs text-slate-500 text-center">
                        <p class="font-semibold text-slate-700 mb-1">Preview Akun (Demo):</p>
                        <p>Guru: <span class="font-mono text-primary bg-indigo-50 px-1 rounded">admin</span> / <span class="font-mono text-primary bg-indigo-50 px-1 rounded">edudigital</span></p>
                        <p class="mt-1">Siswa: <span class="font-mono text-primary bg-indigo-50 px-1 rounded">peserta</span> / <span class="font-mono text-primary bg-indigo-50 px-1 rounded">edudigital</span></p>
                    </div>
                </form>
            </div>
        </div>

        <!-- MAIN DASHBOARD SCREEN -->
        <div id="main-screen" class="hidden flex h-full relative">
            
            <!-- Mobile Sidebar Overlay -->
            <div id="sidebar-overlay" class="fixed inset-0 bg-slate-900/50 z-40 hidden md:hidden backdrop-blur-sm transition-opacity" onclick="toggleSidebar()"></div>

            <!-- SIDEBAR -->
            <div class="w-72 bg-secondary text-slate-300 flex flex-col transition-transform duration-300 shadow-2xl z-50 fixed inset-y-0 left-0 transform -translate-x-full md:relative md:translate-x-0" id="sidebar">
                <div class="p-6 flex items-center justify-between border-b border-slate-800">
                    <div class="flex items-center overflow-hidden">
                        <div id="sidebar-logo" class="mr-3 shrink-0"></div>
                        <h1 class="text-xl font-bold text-white tracking-tight truncate" id="sidebar-school-name">EduAdmin</h1>
                    </div>
                    <button onclick="toggleSidebar()" class="md:hidden text-slate-400 hover:text-white"><i class="fa-solid fa-xmark text-xl"></i></button>
                </div>
                
                <div class="p-6 border-b border-slate-800 bg-slate-800/50">
                    <div class="flex items-center">
                        <div class="w-12 h-12 rounded-full bg-gradient-to-tr from-indigo-500 to-purple-500 flex items-center justify-center text-xl font-bold text-white shadow-md border-2 border-slate-700" id="user-initial">A</div>
                        <div class="ml-4 overflow-hidden">
                            <p class="font-semibold text-white truncate" id="user-name-display">Nama User</p>
                            <span class="inline-flex items-center mt-1 px-2.5 py-0.5 rounded-full text-xs font-medium bg-indigo-500/20 text-indigo-300 border border-indigo-500/30 capitalize" id="user-role-display">Role</span>
                        </div>
                    </div>
                </div>

                <div class="px-4 py-3 text-xs font-semibold text-slate-500 uppercase tracking-wider mt-2">Menu Utama</div>
                <nav class="flex-1 overflow-y-auto px-3 pb-4 space-y-1" id="nav-menu"></nav>

                <div class="p-4 border-t border-slate-800">
                    <button onclick="handleLogout()" class="w-full flex items-center justify-center px-4 py-3 text-red-400 hover:bg-red-500/10 hover:text-red-300 rounded-xl transition font-medium">
                        <i class="fa-solid fa-sign-out-alt mr-2"></i> Keluar Sistem
                    </button>
                </div>
            </div>

            <!-- CONTENT AREA -->
            <div class="flex-1 flex flex-col overflow-hidden bg-slate-50 relative w-full">
                <!-- Header -->
                <header class="bg-white border-b border-slate-200 px-4 md:px-8 py-4 md:py-5 flex justify-between items-center z-10 sticky top-0">
                    <div class="flex items-center overflow-hidden">
                        <button onclick="toggleSidebar()" class="md:hidden mr-4 text-slate-500 hover:text-primary transition focus:outline-none">
                            <i class="fa-solid fa-bars text-xl"></i>
                        </button>
                        <h2 class="text-xl md:text-2xl font-bold text-slate-800 tracking-tight truncate" id="page-title">Dashboard</h2>
                    </div>
                    <div id="loading-indicator" class="hidden items-center px-3 py-1.5 bg-indigo-50 text-primary rounded-full text-xs md:text-sm font-medium border border-indigo-100 whitespace-nowrap ml-2">
                        <i class="fa-solid fa-spinner fa-spin md:mr-2"></i> <span class="hidden md:inline">Menyinkronkan...</span>
                    </div>
                </header>

                <!-- Dynamic Content -->
                <main class="flex-1 scrollable p-4 md:p-8" id="content-area"></main>
            </div>
        </div>
        
    </div>

    <!-- UNIVERSAL MODAL -->
    <div id="universal-modal" class="hidden fixed inset-0 z-[99999999] flex items-center justify-center modal-overlay transition-opacity">
        <div class="bg-white rounded-2xl shadow-2xl w-11/12 max-w-lg flex flex-col max-h-[90vh] fade-in border border-slate-200">
            <div class="px-6 py-5 border-b border-slate-100 flex justify-between items-center bg-slate-50/50 rounded-t-2xl">
                <h3 class="text-xl font-bold text-slate-800" id="modal-title">Form Modal</h3>
                <button onclick="closeModal()" class="text-slate-400 hover:text-slate-600 hover:bg-slate-100 w-8 h-8 rounded-full flex items-center justify-center transition"><i class="fa-solid fa-times text-lg"></i></button>
            </div>
            <div class="p-6 overflow-y-auto flex-1" id="modal-body"></div>
            <div class="px-6 py-5 border-t border-slate-100 flex justify-end gap-3 bg-slate-50 rounded-b-2xl">
                <button onclick="closeModal()" class="px-5 py-2.5 bg-white border border-slate-300 text-slate-700 font-medium rounded-xl hover:bg-slate-50 transition shadow-sm">Batal</button>
                <button onclick="submitModal()" class="px-5 py-2.5 bg-primary text-white font-medium rounded-xl hover:bg-indigo-700 transition shadow-sm shadow-indigo-200" id="modal-submit-btn">Simpan Data</button>
            </div>
        </div>
    </div>

    <!-- WIDGET WAJIB BLOGGER (DISEMBUNYIKAN) -->
    <div id='blogger-mandatory-widgets'>
        <b:section id='main' showaddelement='yes'>
            <b:widget id='Blog1' locked='true' title='Blog Posts' type='Blog' version='2' visible='true'>
                <b:widget-settings>
                    <b:widget-setting name='showDateHeader'>false</b:widget-setting>
                </b:widget-settings>
                <b:includable id='main'>
                    <!-- Kosong, hanya untuk memenuhi syarat struktur Blogger -->
                </b:includable>
            </b:widget>
        </b:section>
    </div>

    <script>
    //<![CDATA[
        // ==========================================
        // KONFIGURASI APLIKASI
        // ==========================================
        // UBAH IS_PREVIEW KE false JIKA AKAN DI-DEPLOY KE BLOGGER
        const IS_PREVIEW = false; 
        const GAS_URL = "https://script.google.com/macros/s/AKfycbwf41X_6gJ8_m6UKhwgawgs29b3mdV_RTJOu8uDsHL_vCBMalx1WT4PgwUIINMaLPpqfg/exec";

        // ==========================================
        // STATE & DUMMY DATA
        // ==========================================
        let currentUser = null;
        let currentMenu = 'dashboard';
        let currentEditingId = null;
        let currentModalContext = '';

        let db = {
            pengaturan: [
                { id: "P1", nama_sekolah: "EduAdmin Portal", logo: "", alamat: "Jl. Pendidikan No. 1, Jakarta", kepsek: "Drs. Budi Setiawan", nip_kepsek: "19700101", guru: "Budi Santoso, S.Pd", nip_guru: "19850101" }
            ],
            users: [
                { id: "U1", username: "admin", password: "edudigital", role: "guru", nama: "Budi Santoso, S.Pd" },
                { id: "U2", username: "peserta", password: "edudigital", role: "siswa", nama: "Andi Saputra", ref_id: "S1" }
            ],
            kelas: [
                { id: "K1", nama: "10 IPA 1" },
                { id: "K2", nama: "10 IPA 2" }
            ],
            siswa: [
                { id: "S1", nis: "1001", nama: "Andi Saputra", id_kelas: "K1", is_bimbingan: "true" },
                { id: "S2", nis: "1002", nama: "Siti Aminah", id_kelas: "K1", is_bimbingan: "false" },
                { id: "S3", nis: "1003", nama: "Bima Arya", id_kelas: "K2", is_bimbingan: "false" }
            ],
            jadwal: [
                { id: "J1", hari: "Senin", jam: "08:00 - 09:30", mapel: "Matematika", id_kelas: "K1" },
                { id: "J2", hari: "Selasa", jam: "10:00 - 11:30", mapel: "Fisika", id_kelas: "K1" },
                { id: "J3", hari: "Senin", jam: "10:00 - 11:30", mapel: "Matematika", id_kelas: "K2" }
            ],
            absen: [
                { id: "ABS_2026-04-12_1001", tanggal: "2026-04-12", nis: "1001", status: "Hadir", ket: "-" }
            ],
            nilai: [
                { id: "NIL_K1_Matematika_Tugas_1001", nis: "1001", mapel: "Matematika", tp: "Memahami Aljabar", jenis: "Tugas", skor: "85" }
            ],
            jurnal: [
                { id: "JR1", tanggal: "2026-04-12", kelas: "K1", mapel: "Matematika", tp: "Fungsi Kuadrat", deskripsi: "Mengerjakan LKS" }
            ],
            bimbingan: [
                { id: "B1", tanggal: "2026-04-10", nis: "1001", jenis: "Akademik", catatan: "Kurang fokus", tindak_lanjut: "Pemanggilan ortu" }
            ],
            dokumen: [
                { id: "D1", nama: "Silabus Matematika Semester 1", link: "https://google.com" },
                { id: "D2", nama: "Buku Pegangan Guru", link: "https://drive.google.com" }
            ]
        };

        const MENUS = {
            guru: [
                { id: 'dashboard', icon: 'fa-chart-pie', title: 'Dashboard' },
                { id: 'kelas', icon: 'fa-chalkboard', title: 'Kelola Kelas' },
                { id: 'siswa', icon: 'fa-users', title: 'Kelola Siswa' },
                { id: 'jadwal', icon: 'fa-calendar-alt', title: 'Kelola Jadwal' },
                { id: 'absen', icon: 'fa-clipboard-user', title: 'Presensi Siswa' },
                { id: 'nilai', icon: 'fa-star', title: 'Manajemen Nilai' },
                { id: 'jurnal', icon: 'fa-book-open', title: 'Jurnal Mengajar' },
                { id: 'bimbingan', icon: 'fa-comments', title: 'Bimbingan Wali' },
                { id: 'dokumen', icon: 'fa-folder-open', title: 'Dokumen Penting' },
                { id: 'pengaturan', icon: 'fa-cog', title: 'Pengaturan Sekolah' }
            ],
            siswa: [
                { id: 'dashboard', icon: 'fa-chart-pie', title: 'Dashboard' },
                { id: 'lihat_absen', icon: 'fa-clipboard-user', title: 'Kehadiran Saya' },
                { id: 'lihat_nilai', icon: 'fa-star', title: 'Nilai Saya' },
                { id: 'lihat_bimbingan', icon: 'fa-comments', title: 'Catatan Bimbingan' }
            ]
        };

        // ==========================================
        // CORE FUNCTIONS
        // ==========================================

        function cleanDate(dateStr) {
            if(!dateStr) return '';
            if(/^\d{4}-\d{2}-\d{2}$/.test(dateStr.trim())) return dateStr.trim();
            try {
                const d = new Date(dateStr);
                if(isNaN(d.getTime())) return dateStr;
                const year = d.getFullYear();
                const month = String(d.getMonth() + 1).padStart(2, '0');
                const day = String(d.getDate()).padStart(2, '0');
                return year + '-' + month + '-' + day;
            } catch(e) {
                if(dateStr.includes('T')) return dateStr.split('T')[0];
                return dateStr;
            }
        }

        function applyPengaturanUI() {
            let set = {nama_sekolah: 'EduAdmin', logo: ''};
            
            // Cek apakah ada data cache pengaturan di localStorage
            const cachedSet = localStorage.getItem('eduAdmin_pengaturan_public');
            
            if (currentUser || IS_PREVIEW) {
                // Jika sedang login atau preview, ambil dari database terbaru
                if (db.pengaturan && db.pengaturan.length > 0) {
                    set = db.pengaturan[0];
                    // Simpan ke cache untuk ditampilkan di halaman login (saat belum login/logout)
                    if (!IS_PREVIEW) localStorage.setItem('eduAdmin_pengaturan_public', JSON.stringify(set));
                }
            } else if (cachedSet) {
                // Jika belum login, gunakan cache agar logo tidak kembali ke default
                set = JSON.parse(cachedSet);
            } else if (db.pengaturan && db.pengaturan.length > 0) {
                // Fallback default bawaan kode
                set = db.pengaturan[0];
            }
            
            // Login Screen UI
            document.getElementById('login-school-name').innerText = set.nama_sekolah;
            if(set.logo) {
                document.getElementById('login-logo-container').innerHTML = '<img src="' + set.logo + '" alt="Logo" class="block w-24 h-24 object-contain mx-auto mb-4 rounded-xl bg-white p-2 shadow border border-slate-100"/>';
            } else {
                document.getElementById('login-logo-container').innerHTML = '<div class="block w-20 h-20 bg-indigo-100 rounded-full flex items-center justify-center mx-auto mb-4"><i class="fa-solid fa-graduation-cap text-4xl text-primary"></i></div>';
            }

            // Sidebar Header UI
            document.getElementById('sidebar-school-name').innerText = set.nama_sekolah;
            if(set.logo) {
                document.getElementById('sidebar-logo').innerHTML = '<img src="' + set.logo + '" alt="Logo" class="block w-10 h-10 object-contain rounded-lg bg-white p-1 shadow-lg shadow-indigo-500/30"/>';
            } else {
                document.getElementById('sidebar-logo').innerHTML = '<div class="w-10 h-10 bg-primary rounded-lg flex items-center justify-center mr-3 shadow-lg shadow-indigo-500/30"><i class="fa-solid fa-graduation-cap text-xl text-white"></i></div>';
            }
            
            if(!IS_PREVIEW) {
                const demoInfo = document.getElementById('demo-info');
                if(demoInfo) demoInfo.style.display = 'none';
            }
        }

        window.addEventListener('DOMContentLoaded', () => {
            const savedUser = sessionStorage.getItem('eduAdmin_user');
            const savedDb = sessionStorage.getItem('eduAdmin_db');
            if (savedUser) {
                currentUser = JSON.parse(savedUser);
                if(savedDb) db = JSON.parse(savedDb);
                applyPengaturanUI();
                initDashboard();
            } else {
                applyPengaturanUI();
            }
        });

        function updateSessionStorage() {
            sessionStorage.setItem('eduAdmin_user', JSON.stringify(currentUser));
            if(!IS_PREVIEW) sessionStorage.setItem('eduAdmin_db', JSON.stringify(db));
        }

        function toggleSidebar() {
            const sidebar = document.getElementById('sidebar');
            const overlay = document.getElementById('sidebar-overlay');
            sidebar.classList.toggle('-translate-x-full');
            overlay.classList.toggle('hidden');
        }

        function showLoading(show) {
            const el = document.getElementById('loading-indicator');
            if (show) { el.classList.remove('hidden'); el.classList.add('flex'); }
            else { el.classList.add('hidden'); el.classList.remove('flex'); }
        }

        async function fetchAPI(action, payload = {}) {
            if (IS_PREVIEW) return new Promise(resolve => setTimeout(() => resolve(mockAPI(action, payload)), 400));
            showLoading(true);
            try {
                const response = await fetch(GAS_URL + "?action=" + action, { method: 'POST', body: JSON.stringify(payload) });
                const result = await response.json();
                showLoading(false);
                return result;
            } catch (error) {
                showLoading(false); alert("Kesalahan koneksi ke server."); return { status: "error", message: "Network Error" };
            }
        }

        function mockAPI(action, payload) {
            if (action === 'login') {
                const user = db.users.find(u => u.username === payload.username && u.password === payload.password);
                if (user) return { status: "success", user: user, data: db };
                return { status: "error", message: "Username atau Password salah!" };
            }
            if (action === 'save') {
                const { table, data, id } = payload;
                if (!db[table]) db[table] = [];
                let savedId = id;
                if (id) {
                    const idx = db[table].findIndex(item => item.id === id);
                    if (idx > -1) db[table][idx] = { ...db[table][idx], ...data };
                } else {
                    savedId = "ID_" + new Date().getTime();
                    db[table].push({ id: savedId, ...data });
                }

                if (table === 'siswa') {
                    const uIdx = db.users.findIndex(u => u.ref_id === savedId);
                    if (uIdx > -1) {
                        db.users[uIdx].username = data.nis;
                        db.users[uIdx].nama = data.nama;
                    } else {
                        db.users.push({ id: "U_" + savedId, username: data.nis, password: "edudigital", role: "siswa", nama: data.nama, ref_id: savedId });
                    }
                }
                return { status: "success" };
            }
            if (action === 'save_bulk') {
                const { table, dataArray } = payload;
                if (!db[table]) db[table] = [];
                dataArray.forEach(item => {
                    const idx = db[table].findIndex(x => x.id === item.id);
                    if (idx > -1) db[table][idx] = { ...db[table][idx], ...item };
                    else db[table].push(item);

                    if (table === 'siswa') {
                        const uIdx = db.users.findIndex(u => u.ref_id === item.id);
                        if (uIdx > -1) {
                            db.users[uIdx].username = item.nis;
                            db.users[uIdx].nama = item.nama;
                        } else {
                            db.users.push({ id: "U_" + item.id, username: item.nis, password: "edudigital", role: "siswa", nama: item.nama, ref_id: item.id });
                        }
                    }
                });
                return { status: "success" };
            }
            if (action === 'delete') {
                db[payload.table] = db[payload.table].filter(item => item.id !== payload.id);
                if (payload.table === 'siswa') {
                    db.users = db.users.filter(u => u.ref_id !== payload.id);
                }
                return { status: "success" };
            }
        }

        async function handleLogin(e) {
            e.preventDefault();
            const u = document.getElementById('username').value;
            const p = document.getElementById('password').value;
            const btn = document.getElementById('btn-login');
            btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin mr-2"></i> Memproses...'; btn.disabled = true;

            const res = await fetchAPI('login', { username: u, password: p });

            if (res.status === 'success') {
                currentUser = res.user; 
                if(!IS_PREVIEW) db = res.data;
                updateSessionStorage(); 
                applyPengaturanUI();
                initDashboard();
            } else {
                const alertBox = document.getElementById('login-alert');
                alertBox.innerText = res.message; alertBox.classList.remove('hidden');
            }
            btn.innerHTML = 'Login ke Portal'; btn.disabled = false;
        }

        function handleLogout() {
            currentUser = null;
            sessionStorage.removeItem('eduAdmin_user');
            sessionStorage.removeItem('eduAdmin_db');
            
            applyPengaturanUI();

            document.getElementById('main-screen').classList.add('hidden');
            document.getElementById('login-screen').classList.remove('hidden');
            document.getElementById('login-form').reset();
            document.getElementById('login-alert').classList.add('hidden');
        }

        function initDashboard() {
            document.getElementById('login-screen').classList.add('hidden');
            document.getElementById('main-screen').classList.remove('hidden');
            document.getElementById('user-initial').innerText = currentUser.nama.charAt(0);
            document.getElementById('user-name-display').innerText = currentUser.nama;
            document.getElementById('user-role-display').innerText = currentUser.role;
            buildNavigation(); loadMenu('dashboard');
        }

        function buildNavigation() {
            const nav = document.getElementById('nav-menu'); nav.innerHTML = '';
            MENUS[currentUser.role].forEach(m => {
                const a = document.createElement('a');
                a.href = '#'; a.id = 'menu-btn-' + m.id;
                a.className = 'flex items-center px-4 py-3 mb-1 text-slate-400 hover:bg-slate-800/50 hover:text-white rounded-xl transition-all duration-200 menu-item group font-medium';
                a.onclick = (e) => { e.preventDefault(); loadMenu(m.id); };
                a.innerHTML = '<i class="fa-solid ' + m.icon + ' w-6 text-center mr-3 text-slate-500 group-hover:text-indigo-400 transition-colors"></i> <span>' + m.title + '</span>';
                nav.appendChild(a);
            });
        }

        function loadMenu(menuId) {
            currentMenu = menuId;
            document.getElementById('page-title').innerText = MENUS[currentUser.role].find(m => m.id === menuId)?.title || 'Dashboard';
            
            document.querySelectorAll('.menu-item').forEach(el => {
                el.classList.remove('bg-indigo-500/10', 'text-white', 'text-indigo-400');
                el.classList.add('text-slate-400');
                el.querySelector('i').classList.remove('text-indigo-400');
                el.querySelector('i').classList.add('text-slate-500');
            });
            
            const activeBtn = document.getElementById('menu-btn-' + menuId);
            if(activeBtn) {
                activeBtn.classList.remove('text-slate-400', 'hover:bg-slate-800/50');
                activeBtn.classList.add('bg-indigo-500/10', 'text-white');
                activeBtn.querySelector('i').classList.remove('text-slate-500');
                activeBtn.querySelector('i').classList.add('text-indigo-400');
            }

            renderContent(menuId);

            if(window.innerWidth < 768) {
                const sidebar = document.getElementById('sidebar');
                const overlay = document.getElementById('sidebar-overlay');
                if(!sidebar.classList.contains('-translate-x-full')) {
                    sidebar.classList.add('-translate-x-full');
                    overlay.classList.add('hidden');
                }
            }
        }

        // ==========================================
        // RENDER VIEWS ROUTER
        // ==========================================
        function renderContent(menuId) {
            const container = document.getElementById('content-area');
            container.className = "flex-1 scrollable p-4 md:p-8 fade-in";

            if (menuId === 'dashboard') container.innerHTML = renderDashboard();
            else if (menuId === 'pengaturan') container.innerHTML = renderPengaturan();
            else if (menuId === 'dokumen') container.innerHTML = renderTableView('dokumen', ['Nama Dokumen', 'Tautan/Link'], ['nama', 'link']);
            else if (menuId === 'kelas') container.innerHTML = renderTableView('kelas', ['Nama Kelas'], ['nama']);
            else if (menuId === 'siswa') container.innerHTML = renderTableView('siswa', ['NIS', 'Nama Siswa', 'Kelas', 'Bimbingan?'], ['nis', 'nama', 'id_kelas', 'is_bimbingan'], true);
            else if (menuId === 'jadwal') container.innerHTML = renderTableView('jadwal', ['Hari', 'Jam', 'Mata Pelajaran', 'Kelas'], ['hari', 'jam', 'mapel', 'id_kelas'], true);
            else if (menuId === 'absen') container.innerHTML = renderAbsenModule();
            else if (menuId === 'nilai') container.innerHTML = renderNilaiModule();
            else if (menuId === 'jurnal') container.innerHTML = renderTableView('jurnal', ['Tanggal', 'Kelas', 'Mata Pelajaran', 'Tujuan Pembelajaran', 'Deskripsi'], ['tanggal', 'kelas', 'mapel', 'tp', 'deskripsi'], true);
            else if (menuId === 'bimbingan') container.innerHTML = renderTableView('bimbingan', ['Tanggal', 'Siswa', 'Jenis', 'Catatan', 'Tindak Lanjut'], ['tanggal', 'nis', 'jenis', 'catatan', 'tindak_lanjut'], true);
            else if (menuId === 'lihat_absen') container.innerHTML = renderSiswaAbsenModule();
            else if (menuId === 'lihat_nilai') container.innerHTML = renderSiswaView('nilai', ['Mapel', 'TP', 'Jenis', 'Skor'], ['mapel', 'tp', 'jenis', 'skor']);
            else if (menuId === 'lihat_bimbingan') container.innerHTML = renderSiswaView('bimbingan', ['Tanggal', 'Jenis', 'Catatan', 'Tindak Lanjut'], ['tanggal', 'jenis', 'catatan', 'tindak_lanjut']);
        }

        function createStatCard(title, value, icon, colorClass, bgIconClass) {
            return '<div class="bg-white rounded-2xl shadow-sm border border-slate-100 p-6 flex items-center hover:shadow-md transition duration-300">' +
                    '<div class="p-4 rounded-xl ' + bgIconClass + ' ' + colorClass + ' mr-5">' +
                        '<i class="fa-solid ' + icon + ' text-2xl"></i>' +
                    '</div>' +
                    '<div>' +
                        '<p class="text-sm font-medium text-slate-500 mb-1">' + title + '</p>' +
                        '<p class="text-3xl font-bold text-slate-800">' + value + '</p>' +
                    '</div>' +
                '</div>';
        }

        function renderDashboard() {
            const hariArray = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
            const hariIni = hariArray[new Date().getDay()];

            if (currentUser.role === 'guru') {
                const jadwalHrIni = db.jadwal.filter(j => j.hari === hariIni);
                
                let jadwalHtml = '<h4 class="font-bold text-slate-700 mt-8 mb-4 border-b border-slate-200 pb-2 flex items-center"><i class="fa-solid fa-calendar-day mr-2 text-primary"></i> Jadwal Pelajaran Hari Ini (' + hariIni + ')</h4>';
                if(jadwalHrIni.length === 0) jadwalHtml += '<p class="text-sm text-slate-500 italic bg-slate-50 p-4 rounded-xl border border-slate-200">Tidak ada jadwal kelas pada hari ini.</p>';
                else {
                    jadwalHtml += '<div class="grid grid-cols-1 md:grid-cols-3 gap-4">';
                    jadwalHrIni.forEach(j => {
                        const kls = db.kelas.find(k=>k.id===j.id_kelas)?.nama;
                        jadwalHtml += '<div class="bg-gradient-to-br from-indigo-50 to-white border border-indigo-100 p-5 rounded-xl shadow-sm hover:shadow-md transition">' +
                            '<p class="text-xs font-bold text-indigo-500 uppercase tracking-widest mb-1"><i class="fa-regular fa-clock mr-1"></i> ' + j.jam + '</p>' +
                            '<p class="font-bold text-slate-800 text-lg">' + j.mapel + '</p>' +
                            '<p class="text-sm text-slate-600 font-medium mt-1"><i class="fa-solid fa-door-open mr-1 text-slate-400"></i> Kelas: ' + kls + '</p>' +
                        '</div>';
                    });
                    jadwalHtml += '</div>';
                }

                return '<div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 md:gap-6 mb-6 md:mb-8">' +
                    createStatCard('Total Siswa', db.siswa.length, 'fa-users', 'text-blue-600', 'bg-blue-50') +
                    createStatCard('Total Kelas', db.kelas.length, 'fa-chalkboard', 'text-emerald-600', 'bg-emerald-50') +
                    createStatCard('Jurnal Terisi', db.jurnal.length, 'fa-book-open', 'text-indigo-600', 'bg-indigo-50') +
                    createStatCard('Kasus Bimbingan', db.bimbingan.length, 'fa-comments', 'text-rose-600', 'bg-rose-50') +
                '</div>' +
                '<div class="bg-white p-5 md:p-8 rounded-2xl shadow-sm border border-slate-100 flex items-start gap-4 md:gap-6 flex-col md:flex-row">' +
                    '<div class="p-4 bg-indigo-50 rounded-full text-primary hidden md:block"><i class="fa-solid fa-hand-sparkles text-3xl"></i></div>' +
                    '<div class="w-full">' +
                        '<h3 class="font-bold text-xl md:text-2xl mb-2 text-slate-800">Selamat Datang, ' + currentUser.nama + '</h3>' +
                        '<p class="text-slate-600 text-sm md:text-base leading-relaxed">Sistem administrasi kini saling terhubung. Pastikan Anda mengatur <b class="text-slate-800">Jadwal</b> kelas terlebih dahulu. Modul <b class="text-slate-800">Input Nilai</b> dan <b class="text-slate-800">Jurnal</b> akan secara otomatis mendeteksi Mata Pelajaran berdasarkan Kelas yang dipilih sesuai dengan Jadwal yang telah dibuat.</p>' +
                        jadwalHtml +
                    '</div>' +
                '</div>';
            } else {
                const student = db.siswa.find(s => s.id === currentUser.ref_id);
                if(!student) return '<div class="p-4 bg-red-100 text-red-700 rounded-xl">Data siswa tidak valid.</div>';
                
                const nis = student.nis;
                const absensiKu = db.absen.filter(a => a.nis === nis);
                let hadir = 0; let totalAbsen = 0;
                absensiKu.forEach(a => {
                    if(['Hadir', 'Sakit', 'Izin', 'Alpa'].includes(a.status)) totalAbsen++;
                    if(a.status === 'Hadir') hadir++;
                });
                const persenHadir = totalAbsen > 0 ? Math.round((hadir/totalAbsen)*100) : 0;
                
                const nilaiKu = db.nilai.filter(n => n.nis === nis);
                let totalSkor = 0;
                nilaiKu.forEach(n => totalSkor += (parseFloat(n.skor)||0));
                const rataRata = nilaiKu.length > 0 ? (totalSkor / nilaiKu.length).toFixed(1) : 0;

                const jadwalHrIni = db.jadwal.filter(j => j.hari === hariIni && j.id_kelas === student.id_kelas);
                let jadwalHtml = '<h4 class="font-bold text-slate-700 mt-6 mb-4 border-b border-slate-200 pb-2"><i class="fa-solid fa-calendar-day mr-2 text-primary"></i> Jadwal Anda Hari Ini (' + hariIni + ')</h4>';
                
                if(jadwalHrIni.length === 0) jadwalHtml += '<p class="text-sm text-slate-500 italic bg-white p-4 rounded-xl border border-slate-200 shadow-sm">Bebas kelas! Tidak ada jadwal hari ini.</p>';
                else {
                    jadwalHtml += '<div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">';
                    jadwalHrIni.forEach(j => {
                        jadwalHtml += '<div class="bg-white border border-slate-200 p-5 rounded-xl shadow-sm hover:shadow-md transition">' +
                            '<p class="text-xs font-bold text-indigo-500 uppercase tracking-widest mb-1"><i class="fa-regular fa-clock mr-1"></i> ' + j.jam + '</p>' +
                            '<p class="font-bold text-slate-800 text-lg">' + j.mapel + '</p>' +
                        '</div>';
                    });
                    jadwalHtml += '</div>';
                }

                return '<div class="bg-gradient-to-r from-primary to-purple-600 p-6 md:p-8 rounded-2xl shadow-lg text-white mb-6 flex items-center justify-between">' +
                    '<div>' +
                        '<h3 class="font-bold text-2xl md:text-3xl mb-1">Halo, ' + currentUser.nama + '!</h3>' +
                        '<p class="text-indigo-100 text-sm md:text-lg">Selamat datang di Portal Akademik Siswa. Tetap semangat belajar hari ini.</p>' +
                    '</div>' +
                    '<i class="fa-solid fa-rocket text-5xl text-white/20 hidden md:block"></i>' +
                '</div>' +
                '<div class="grid grid-cols-1 md:grid-cols-2 gap-4 md:gap-6 mb-6">' +
                    createStatCard('Persentase Kehadiran', persenHadir + '%', 'fa-user-check', persenHadir >= 80 ? 'text-emerald-600' : 'text-rose-600', persenHadir >= 80 ? 'bg-emerald-50' : 'bg-rose-50') +
                    createStatCard('Rata-Rata Nilai Keseluruhan', rataRata, 'fa-star', rataRata >= 75 ? 'text-blue-600' : 'text-amber-500', rataRata >= 75 ? 'bg-blue-50' : 'bg-amber-50') +
                '</div>' +
                jadwalHtml;
            }
        }

        function renderTableView(tableId, headers, keys, needsRel = false) {
            let data = db[tableId] || [];
            if(tableId === 'bimbingan') data = data.filter(d => db.siswa.find(s => s.nis === d.nis && (s.is_bimbingan === "true" || s.is_bimbingan === true)));

            let btnExtra = '';
            let infoText = '';
            if(tableId === 'siswa') {
                btnExtra = '<button onclick="downloadFormatExcel()" class="bg-emerald-50 text-emerald-600 px-4 py-2.5 rounded-xl border border-emerald-200 hover:bg-emerald-100 text-sm font-medium transition mr-2" title="Unduh Format Excel"><i class="fa-solid fa-download"></i></button>' +
                            '<button onclick="triggerImportExcel()" class="bg-emerald-600 text-white px-4 py-2.5 rounded-xl shadow-sm hover:bg-emerald-700 text-sm font-medium transition mr-2 flex items-center"><i class="fa-solid fa-file-import mr-2"></i> Import Siswa</button>';
                infoText = '<div class="bg-blue-50 border border-blue-100 text-blue-700 p-3 rounded-xl mb-4 text-sm font-medium flex items-center"><i class="fa-solid fa-circle-info mr-2"></i> Akun login siswa otomatis dibuat! Username: <b class="mx-1">NIS</b> | Password: <b class="ml-1">edudigital</b></div>';
            }
            if(tableId === 'jurnal' || tableId === 'bimbingan') {
                btnExtra = '<button onclick="cetakLaporan(\'tabel-print-' + tableId + '\', \'Laporan ' + tableId.toUpperCase() + '\')" class="bg-slate-800 text-white px-4 py-2.5 rounded-xl shadow-sm hover:bg-slate-700 text-sm font-medium transition mr-2 flex items-center"><i class="fa-solid fa-print mr-2"></i> Cetak</button>';
            }

            let html = infoText + 
                '<div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">' +
                    '<div class="px-6 py-5 border-b border-slate-200 flex flex-col sm:flex-row justify-between items-center bg-white gap-4">' +
                        '<h3 class="font-bold text-slate-800 text-lg capitalize">Master Data ' + tableId + '</h3>' +
                        '<div class="flex">' +
                            btnExtra +
                            '<button onclick="openModal(\'' + tableId + '\')" class="bg-primary text-white px-5 py-2.5 rounded-xl shadow-sm hover:bg-indigo-700 hover:shadow text-sm font-medium transition flex items-center">' +
                                '<i class="fa-solid fa-plus mr-2"></i> Tambah Data Baru' +
                            '</button>' +
                        '</div>' +
                    '</div>' +
                    '<div class="overflow-x-auto" id="tabel-print-' + tableId + '">' +
                        '<table class="w-full text-left border-collapse whitespace-nowrap">' +
                            '<thead>' +
                                '<tr class="bg-slate-50 text-slate-500 text-xs uppercase tracking-wider font-semibold border-b border-slate-200">' +
                                    '<th class="px-6 py-4">No</th>' +
                                    headers.map(h => '<th class="px-6 py-4">' + h + '</th>').join('') +
                                    '<th class="px-6 py-4 text-center">Aksi</th>' +
                                '</tr>' +
                            '</thead>' +
                            '<tbody class="text-slate-700 text-sm divide-y divide-slate-100">';

            if (data.length === 0) html += '<tr><td colspan="' + (headers.length + 2) + '" class="px-6 py-8 text-center text-slate-400 italic">Belum ada data tersedia.</td></tr>';
            else {
                data.forEach((row, idx) => {
                    html += '<tr class="hover:bg-slate-50/80 transition"><td class="px-6 py-4 text-slate-500">' + (idx + 1) + '</td>';
                    keys.forEach(k => {
                        let val = row[k];
                        let valText = val; 
                        let isHtmlColumn = false;
                        
                        if (k === 'id_kelas' && needsRel) val = db.kelas.find(x => x.id === val)?.nama || val;
                        if (k === 'nis' && needsRel) val = db.siswa.find(x => x.nis === val)?.nama || val;
                        if (k === 'kelas' && needsRel) val = db.kelas.find(x => x.id === val)?.nama || val;
                        
                        if (k === 'tanggal') {
                            val = cleanDate(val);
                            valText = val;
                        } else if (k === 'link') {
                            val = '<a href="' + val + '" target="_blank" class="text-blue-600 hover:text-blue-800 hover:underline font-medium flex items-center"><i class="fa-solid fa-link mr-1.5 text-xs"></i> Buka Tautan</a>';
                            isHtmlColumn = true;
                        } else if (k === 'is_bimbingan') {
                            valText = (val === "true" || val === true) ? 'Ya' : 'Tidak';
                            val = (val === "true" || val === true) ? '<span class="px-2.5 py-1 bg-rose-100 text-rose-700 rounded-full text-xs font-semibold">Ya</span>' : '<span class="px-2.5 py-1 bg-slate-100 text-slate-600 rounded-full text-xs font-semibold">Tidak</span>';
                            isHtmlColumn = true;
                        }
                        
                        if(isHtmlColumn) {
                            html += '<td class="px-6 py-4 max-w-[200px] truncate">' + (val || '-') + '</td>';
                        } else {
                            html += '<td class="px-6 py-4 max-w-[200px] truncate" title="' + (valText || '-') + '">' + (val || '-') + '</td>';
                        }
                    });
                    html += '<td class="px-6 py-4 text-center no-print">' +
                            '<button onclick="openModal(\'' + tableId + '\', \'' + row.id + '\')" class="w-8 h-8 rounded-lg bg-indigo-50 text-indigo-600 hover:bg-indigo-100 hover:text-indigo-800 transition mx-1" title="Edit"><i class="fa-solid fa-pen-to-square text-sm"></i></button>' +
                            '<button onclick="deleteData(\'' + tableId + '\', \'' + row.id + '\')" class="w-8 h-8 rounded-lg bg-red-50 text-red-600 hover:bg-red-100 hover:text-red-800 transition mx-1" title="Hapus"><i class="fa-solid fa-trash-can text-sm"></i></button>' +
                        '</td></tr>';
                });
            }
            html += '</tbody></table></div></div>';
            return html;
        }

        function renderSiswaView(tableId, headers, keys) {
            const studentNis = db.siswa.find(s => s.id === currentUser.ref_id)?.nis;
            let data = (db[tableId] || []).filter(item => item.nis === studentNis);
            let html = '<div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden"><div class="px-6 py-5 border-b border-slate-200 bg-slate-50"><h3 class="font-bold text-slate-800 capitalize text-lg">Data ' + tableId + '</h3></div>' +
                '<div class="overflow-x-auto"><table class="w-full text-left border-collapse whitespace-nowrap"><thead><tr class="bg-white text-slate-500 text-xs uppercase tracking-wider font-semibold border-b border-slate-200">' +
                '<th class="px-6 py-4">No</th>' + headers.map(h => '<th class="px-6 py-4">' + h + '</th>').join('') + '</tr></thead><tbody class="text-slate-700 text-sm divide-y divide-slate-100">';
            
            if (data.length === 0) html += '<tr><td colspan="' + (headers.length + 1) + '" class="px-6 py-8 text-center text-slate-400 italic">Belum ada data tersedia.</td></tr>';
            else data.forEach((row, idx) => {
                html += '<tr class="hover:bg-slate-50 transition"><td class="px-6 py-4">' + (idx + 1) + '</td>';
                keys.forEach(k => {
                    let val = row[k];
                    if(k === 'tanggal') val = cleanDate(val);
                    html += '<td class="px-6 py-4">' + (val || '-') + '</td>';
                });
                html += '</tr>';
            });
            html += '</tbody></table></div></div>';
            return html;
        }

        // --- MODULE ABSEN SISWA (FILTER) ---
        function renderSiswaAbsenModule() {
            return '<div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 mb-6 flex flex-col md:flex-row gap-5 items-end no-print">' +
                '<div class="w-full md:w-1/3">' +
                    '<label class="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Jenis Filter</label>' +
                    '<select id="siswa-abs-tipe" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary outline-none bg-slate-50" onchange="toggleSiswaAbsenInput()">' +
                        '<option value="semua">Semua Riwayat</option>' +
                        '<option value="harian">Filter Harian</option>' +
                        '<option value="bulanan">Filter Bulanan</option>' +
                        '<option value="semester">Filter Semester</option>' +
                    '</select>' +
                '</div>' +
                '<div class="w-full md:w-2/3" id="siswa-abs-waktu-container" style="display:none;"></div>' +
            '</div>' +
            '<div id="siswa-abs-container"></div>';
        }

        function toggleSiswaAbsenInput() {
            const tipe = document.getElementById('siswa-abs-tipe').value;
            const container = document.getElementById('siswa-abs-waktu-container');
            
            if(tipe === 'semua') {
                container.style.display = 'none';
                container.innerHTML = '';
            } else {
                container.style.display = 'block';
                if(tipe === 'semester') {
                    container.innerHTML = '<label class="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Tahun &amp; Semester</label>' +
                        '<div class="flex gap-2">' +
                            '<input type="number" id="siswa-abs-tahun" class="w-1/2 p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary outline-none bg-slate-50" value="' + new Date().getFullYear() + '" onchange="loadSiswaAbsenList()" title="Tahun Ajaran"/>' +
                            '<select id="siswa-abs-smt" class="w-1/2 p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary outline-none bg-slate-50" onchange="loadSiswaAbsenList()">' +
                                '<option value="ganjil">Ganjil (Jul-Des)</option>' +
                                '<option value="genap">Genap (Jan-Jun)</option>' +
                            '</select>' +
                        '</div>';
                } else {
                    container.innerHTML = '<label class="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Waktu / Tanggal</label>' +
                        '<input type="' + (tipe === 'harian' ? 'date' : 'month') + '" id="siswa-abs-waktu" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary outline-none bg-slate-50" onchange="loadSiswaAbsenList()"/>';
                }
            }
            loadSiswaAbsenList();
        }

        function loadSiswaAbsenList() {
            const studentNis = db.siswa.find(s => s.id === currentUser.ref_id)?.nis;
            let data = (db.absen || []).filter(item => item.nis === studentNis);
            
            const tipeEl = document.getElementById('siswa-abs-tipe');
            if(!tipeEl) return;

            const tipe = tipeEl.value;

            if(tipe === 'harian') {
                const waktu = document.getElementById('siswa-abs-waktu')?.value;
                if(waktu) data = data.filter(d => cleanDate(d.tanggal) === waktu);
            } else if(tipe === 'bulanan') {
                const waktu = document.getElementById('siswa-abs-waktu')?.value;
                if(waktu) data = data.filter(d => cleanDate(d.tanggal).startsWith(waktu));
            } else if(tipe === 'semester') {
                const tahun = document.getElementById('siswa-abs-tahun')?.value;
                const smt = document.getElementById('siswa-abs-smt')?.value;
                if(tahun && smt) {
                    data = data.filter(d => {
                        const dStr = cleanDate(d.tanggal);
                        if(!dStr) return false;
                        const dateObj = new Date(dStr);
                        const month = dateObj.getMonth() + 1;
                        const y = dateObj.getFullYear();
                        if (smt === 'ganjil') return y == tahun && month >= 7 && month <= 12;
                        if (smt === 'genap') return y == tahun && month >= 1 && month <= 6;
                        return false;
                    });
                }
            }

            let html = '<div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden"><div class="px-6 py-5 border-b border-slate-200 bg-slate-50"><h3 class="font-bold text-slate-800 capitalize text-lg">Riwayat Kehadiran</h3></div>' +
                '<div class="overflow-x-auto"><table class="w-full text-left border-collapse whitespace-nowrap"><thead><tr class="bg-white text-slate-500 text-xs uppercase tracking-wider font-semibold border-b border-slate-200">' +
                '<th class="px-6 py-4">No</th><th class="px-6 py-4">Tanggal</th><th class="px-6 py-4 text-center">Status</th><th class="px-6 py-4">Keterangan</th></tr></thead><tbody class="text-slate-700 text-sm divide-y divide-slate-100">';
            
            if (data.length === 0) html += '<tr><td colspan="4" class="px-6 py-8 text-center text-slate-400 italic">Belum ada catatan kehadiran pada periode ini.</td></tr>';
            else {
                data.forEach((row, idx) => {
                    let colorStatus = 'text-slate-500';
                    if(row.status === 'Hadir') colorStatus = 'text-emerald-600 bg-emerald-50 px-2 py-1 rounded-md';
                    else if(row.status === 'Alpa') colorStatus = 'text-rose-600 bg-rose-50 px-2 py-1 rounded-md';
                    else colorStatus = 'text-amber-600 bg-amber-50 px-2 py-1 rounded-md';

                    html += '<tr class="hover:bg-slate-50 transition">' +
                        '<td class="px-6 py-4">' + (idx + 1) + '</td>' +
                        '<td class="px-6 py-4">' + cleanDate(row.tanggal) + '</td>' +
                        '<td class="px-6 py-4 text-center font-bold"><span class="' + colorStatus + '">' + row.status + '</span></td>' +
                        '<td class="px-6 py-4">' + (row.ket || '-') + '</td>' +
                    '</tr>';
                });
            }
            html += '</tbody></table></div></div>';
            document.getElementById('siswa-abs-container').innerHTML = html;
        }

        function switchTab(tabId, btnEl) {
            document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
            document.getElementById(tabId).classList.remove('hidden');
            document.getElementById(tabId).classList.add('fade-in');
            
            document.querySelectorAll('.tab-btn').forEach(btn => {
                btn.classList.remove('text-primary', 'border-primary', 'bg-indigo-50/50');
                btn.classList.add('border-transparent', 'text-slate-500');
            });
            btnEl.classList.add('text-primary', 'border-primary', 'bg-indigo-50/50');
            btnEl.classList.remove('border-transparent', 'text-slate-500');
        }

        function renderAbsenModule() {
            return '<div class="mb-6 flex space-x-1 border-b border-slate-200 no-print">' +
                '<button class="px-5 py-3 text-sm font-semibold text-primary border-b-2 border-primary bg-indigo-50/50 rounded-t-lg transition tab-btn" onclick="switchTab(\'tab-absen-input\', this)">Input Absen Kelas</button>' +
                '<button class="px-5 py-3 text-sm font-semibold text-slate-500 border-b-2 border-transparent hover:text-slate-700 hover:bg-slate-50 rounded-t-lg transition tab-btn" onclick="switchTab(\'tab-absen-rekap\', this); loadRekapAbsen();">Laporan Rekapitulasi</button>' +
            '</div>' +
            '<div id="tab-absen-input" class="tab-content block">' +
                '<div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 mb-6 flex flex-col md:flex-row gap-5 items-end no-print">' +
                    '<div class="w-full md:w-1/3">' +
                        '<label class="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Tanggal Presensi</label>' +
                        '<input type="date" id="abs-date" value="' + new Date().toISOString().split('T')[0] + '" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary focus:border-transparent outline-none bg-slate-50" onchange="loadAbsenList()"/>' +
                    '</div>' +
                    '<div class="w-full md:w-1/3">' +
                        '<label class="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Pilih Kelas</label>' +
                        '<select id="abs-kelas" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary focus:border-transparent outline-none bg-slate-50" onchange="loadAbsenList()">' +
                            '<option value="">-- Pilih Kelas --</option>' +
                            db.kelas.map(k => '<option value="' + k.id + '">' + k.nama + '</option>').join('') +
                        '</select>' +
                    '</div>' +
                '</div>' +
                '<div id="abs-list-container"></div>' +
            '</div>' +
            '<div id="tab-absen-rekap" class="tab-content hidden">' +
                '<div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 mb-6 flex flex-col md:flex-row justify-between items-end gap-5 no-print">' +
                    '<div class="grid grid-cols-1 md:grid-cols-3 gap-5 flex-1">' +
                        '<div>' +
                            '<label class="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Jenis Filter</label>' +
                            '<select id="rek-abs-tipe" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary outline-none bg-slate-50" onchange="toggleRekapAbsenInput()">' +
                                '<option value="harian">Laporan Harian</option>' +
                                '<option value="bulanan">Laporan Bulanan</option>' +
                                '<option value="semester">Laporan Semester</option>' +
                            '</select>' +
                        '</div>' +
                        '<div>' +
                            '<label class="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Waktu / Periode</label>' +
                            '<div id="rek-abs-waktu-container">' +
                                '<input type="date" id="rek-abs-waktu" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary outline-none bg-slate-50" onchange="loadRekapAbsen()"/>' +
                            '</div>' +
                        '</div>' +
                        '<div>' +
                            '<label class="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Filter Kelas</label>' +
                            '<select id="rek-abs-kelas" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary outline-none bg-slate-50" onchange="loadRekapAbsen()">' +
                                '<option value="">-- Semua Kelas --</option>' +
                                db.kelas.map(k => '<option value="' + k.id + '">' + k.nama + '</option>').join('') +
                            '</select>' +
                        '</div>' +
                    '</div>' +
                    '<button onclick="cetakRekapAbsen()" class="bg-slate-800 text-white font-medium px-6 py-2.5 rounded-xl shadow-md hover:bg-slate-700 transition flex items-center"><i class="fa-solid fa-print mr-2"></i> Cetak</button>' +
                '</div>' +
                '<div id="rek-abs-container"></div>' +
            '</div>';
        }

        function loadAbsenList() {
            const date = document.getElementById('abs-date').value;
            const id_kelas = document.getElementById('abs-kelas').value;
            if(!date || !id_kelas) { document.getElementById('abs-list-container').innerHTML = ''; return; }

            const siswa = db.siswa.filter(s => s.id_kelas === id_kelas);
            let html = '<div class="bg-white shadow-sm border border-slate-200 rounded-2xl overflow-hidden mb-5"><table class="w-full text-left border-collapse text-sm">' +
                '<thead class="bg-slate-50 border-b border-slate-200 text-slate-500 text-xs uppercase tracking-wider"><th class="p-4 pl-6">Nama Siswa</th><th class="p-4 w-56 text-center">Kehadiran</th><th class="p-4 pr-6">Catatan Keterangan</th></thead><tbody class="divide-y divide-slate-100">';
            
            if(siswa.length === 0) html += '<tr><td colspan="3" class="p-6 text-center text-slate-400 italic">Tidak ada siswa di kelas ini.</td></tr>';
            
            siswa.forEach(s => {
                const existing = db.absen.find(a => a.nis === s.nis && cleanDate(a.tanggal) === date) || {status: 'Hadir', ket: ''};
                html += '<tr class="hover:bg-slate-50 transition" data-nis="' + s.nis + '">' +
                    '<td class="p-4 pl-6 font-medium text-slate-700">' + s.nama + ' <span class="text-slate-400 font-normal text-xs ml-1">(' + s.nis + ')</span></td>' +
                    '<td class="p-4">' +
                        '<div class="flex gap-2 justify-center">' +
                            '<label class="cursor-pointer" title="Hadir"><input type="radio" name="abs_' + s.nis + '" value="Hadir" class="peer hidden abs-radio" ' + (existing.status==='Hadir'||!existing.status?'checked="checked"':'') + '/><div class="w-8 h-8 rounded-full flex items-center justify-center font-bold text-xs border border-slate-300 peer-checked:bg-emerald-500 peer-checked:text-white peer-checked:border-emerald-500 text-slate-400 transition hover:bg-slate-100">H</div></label>' +
                            '<label class="cursor-pointer" title="Sakit"><input type="radio" name="abs_' + s.nis + '" value="Sakit" class="peer hidden abs-radio" ' + (existing.status==='Sakit'?'checked="checked"':'') + '/><div class="w-8 h-8 rounded-full flex items-center justify-center font-bold text-xs border border-slate-300 peer-checked:bg-amber-500 peer-checked:text-white peer-checked:border-amber-500 text-slate-400 transition hover:bg-slate-100">S</div></label>' +
                            '<label class="cursor-pointer" title="Izin"><input type="radio" name="abs_' + s.nis + '" value="Izin" class="peer hidden abs-radio" ' + (existing.status==='Izin'?'checked="checked"':'') + '/><div class="w-8 h-8 rounded-full flex items-center justify-center font-bold text-xs border border-slate-300 peer-checked:bg-blue-500 peer-checked:text-white peer-checked:border-blue-500 text-slate-400 transition hover:bg-slate-100">I</div></label>' +
                            '<label class="cursor-pointer" title="Alpa"><input type="radio" name="abs_' + s.nis + '" value="Alpa" class="peer hidden abs-radio" ' + (existing.status==='Alpa'?'checked="checked"':'') + '/><div class="w-8 h-8 rounded-full flex items-center justify-center font-bold text-xs border border-slate-300 peer-checked:bg-rose-600 peer-checked:text-white peer-checked:border-rose-600 text-slate-400 transition hover:bg-slate-100">A</div></label>' +
                        '</div>' +
                    '</td>' +
                    '<td class="p-4 pr-6"><input type="text" class="abs-ket w-full p-2.5 border border-slate-300 rounded-xl focus:ring-primary focus:border-primary outline-none text-sm bg-white" value="' + (existing.ket||"") + '" placeholder="Opsional..."/></td>' +
                '</tr>';
            });
            html += '</tbody></table></div>';
            if(siswa.length > 0) html += '<button onclick="saveBulkAbsen()" class="bg-primary text-white font-medium px-6 py-3 rounded-xl shadow-md shadow-indigo-200 hover:bg-indigo-700 transition flex items-center justify-center w-full md:w-auto"><i class="fa-solid fa-save mr-2"></i> Simpan Data Kehadiran</button>';
            document.getElementById('abs-list-container').innerHTML = html;
        }

        async function saveBulkAbsen() {
            const date = document.getElementById('abs-date').value;
            const rows = document.querySelectorAll('#abs-list-container tbody tr');
            let dataArray = [];
            rows.forEach(tr => {
                if(!tr.hasAttribute('data-nis')) return;
                const nis = tr.getAttribute('data-nis');
                const radio = tr.querySelector('input[type="radio"]:checked');
                const status = radio ? radio.value : 'Hadir';
                const ket = tr.querySelector('.abs-ket').value;
                dataArray.push({ id: 'ABS_' + date + '_' + nis, tanggal: date, nis: nis, status: status, ket: ket });
            });

            const res = await fetchAPI('save_bulk', { table: 'absen', dataArray: dataArray });
            if(res.status === 'success') {
                if(!IS_PREVIEW) { 
                    const fresh = await fetchAPI('login', {username: currentUser.username, password: currentUser.password}); 
                    db = fresh.data;
                    updateSessionStorage();
                }
                alert("Data kehadiran berhasil disimpan!");
            }
        }

        function toggleRekapAbsenInput() {
            const tipe = document.getElementById('rek-abs-tipe').value;
            const container = document.getElementById('rek-abs-waktu-container');
            
            if(tipe === 'semester') {
                container.innerHTML = '<div class="flex gap-2">' +
                        '<input type="number" id="rek-abs-tahun" class="w-1/2 p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary outline-none bg-slate-50" value="' + new Date().getFullYear() + '" onchange="loadRekapAbsen()" title="Tahun Ajaran"/>' +
                        '<select id="rek-abs-smt" class="w-1/2 p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary outline-none bg-slate-50" onchange="loadRekapAbsen()">' +
                            '<option value="ganjil">Ganjil (Jul-Des)</option>' +
                            '<option value="genap">Genap (Jan-Jun)</option>' +
                        '</select>' +
                    '</div>';
            } else {
                container.innerHTML = '<input type="' + (tipe === 'harian' ? 'date' : 'month') + '" id="rek-abs-waktu" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary outline-none bg-slate-50" onchange="loadRekapAbsen()"/>';
            }
            loadRekapAbsen();
        }

        function loadRekapAbsen() {
            const tipe = document.getElementById('rek-abs-tipe').value;
            const id_kelas = document.getElementById('rek-abs-kelas').value;
            
            let waktu = '';
            let tahun = '', smt = '';

            if(tipe === 'semester') {
                tahun = document.getElementById('rek-abs-tahun')?.value;
                smt = document.getElementById('rek-abs-smt')?.value;
                if(!tahun) return;
            } else {
                waktu = document.getElementById('rek-abs-waktu')?.value;
                if(!waktu) return;
            }

            let siswaList = db.siswa;
            if(id_kelas) siswaList = siswaList.filter(s => s.id_kelas === id_kelas);

            let html = '<div class="bg-white shadow-sm border border-slate-200 rounded-2xl overflow-hidden"><table class="w-full text-left border-collapse text-sm">';
            
            if(tipe === 'harian') {
                html += '<thead class="bg-slate-50 border-b border-slate-200 text-slate-500 text-xs uppercase tracking-wider"><th class="p-4 pl-6">Nama Siswa</th><th class="p-4">Kelas</th><th class="p-4 pr-6">Status</th></thead><tbody class="divide-y divide-slate-100">';
                siswaList.forEach(s => {
                    const abs = db.absen.find(a => a.nis === s.nis && cleanDate(a.tanggal) === waktu);
                    const kls = db.kelas.find(k => k.id === s.id_kelas)?.nama;
                    let colorStatus = 'text-slate-400';
                    if(abs) {
                        if(abs.status === 'Hadir') colorStatus = 'text-emerald-600 bg-emerald-50 px-2 py-1 rounded-md';
                        else if(abs.status === 'Alpa') colorStatus = 'text-rose-600 bg-rose-50 px-2 py-1 rounded-md';
                        else colorStatus = 'text-amber-600 bg-amber-50 px-2 py-1 rounded-md';
                    }
                    html += '<tr class="hover:bg-slate-50 transition"><td class="p-4 pl-6 font-medium text-slate-700">' + s.nama + '</td><td class="p-4 text-slate-500">' + kls + '</td><td class="p-4 pr-6 font-semibold"><span class="' + colorStatus + '">' + (abs ? abs.status : 'Belum Input') + '</span></td></tr>';
                });
            } else {
                html += '<thead class="bg-slate-50 border-b border-slate-200 text-slate-500 text-xs uppercase tracking-wider"><th class="p-4 pl-6">Nama Siswa</th><th class="p-4 text-center">Hadir</th><th class="p-4 text-center">Sakit</th><th class="p-4 text-center">Izin</th><th class="p-4 text-center">Alpa</th><th class="p-4 text-center pr-6 bg-indigo-50/50 text-primary">% Kehadiran</th></thead><tbody class="divide-y divide-slate-100">';
                siswaList.forEach(s => {
                    let h=0, sk=0, i=0, a=0;
                    db.absen.filter(abs => {
                        if(abs.nis !== s.nis) return false;
                        const dStr = cleanDate(abs.tanggal);
                        if(!dStr) return false;
                        
                        if (tipe === 'bulanan') {
                            return dStr.startsWith(waktu);
                        } else if (tipe === 'semester') {
                            const dateObj = new Date(dStr);
                            const month = dateObj.getMonth() + 1;
                            const y = dateObj.getFullYear();
                            if (smt === 'ganjil') return y == tahun && month >= 7 && month <= 12;
                            if (smt === 'genap') return y == tahun && month >= 1 && month <= 6;
                        }
                        return false;
                    }).forEach(abs => {
                        if(abs.status==='Hadir') h++; else if(abs.status==='Sakit') sk++; else if(abs.status==='Izin') i++; else if(abs.status==='Alpa') a++;
                    });
                    
                    const totalAktif = h + sk + i + a;
                    const persen = totalAktif > 0 ? Math.round((h / totalAktif) * 100) : 0;
                    const persenWarna = persen < 80 ? 'text-rose-600' : 'text-emerald-600';
                    html += '<tr class="hover:bg-slate-50 transition"><td class="p-4 pl-6 font-medium text-slate-700">' + s.nama + '</td><td class="p-4 text-center text-emerald-600 font-bold">' + h + '</td><td class="p-4 text-center text-amber-500">' + sk + '</td><td class="p-4 text-center text-blue-500">' + i + '</td><td class="p-4 text-center text-rose-600 font-bold">' + a + '</td><td class="p-4 text-center pr-6 font-bold bg-indigo-50/30 ' + persenWarna + '">' + persen + '%</td></tr>';
                });
            }
            html += '</tbody></table></div>';
            document.getElementById('rek-abs-container').innerHTML = html;
        }

        // --- MODULE NILAI ---
        function renderNilaiModule() {
            return '<div class="mb-6 flex space-x-1 border-b border-slate-200 no-print">' +
                '<button class="px-5 py-3 text-sm font-semibold text-primary border-b-2 border-primary bg-indigo-50/50 rounded-t-lg transition tab-btn" onclick="switchTab(\'tab-nilai-input\', this)">Input Nilai Masal</button>' +
                '<button class="px-5 py-3 text-sm font-semibold text-slate-500 border-b-2 border-transparent hover:text-slate-700 hover:bg-slate-50 rounded-t-lg transition tab-btn" onclick="switchTab(\'tab-nilai-rekap\', this); loadRekapNilai();">Rekapitulasi Nilai</button>' +
            '</div>' +
            '<div id="tab-nilai-input" class="tab-content block">' +
                '<div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 mb-6 grid grid-cols-1 md:grid-cols-4 gap-5 no-print">' +
                    '<div>' +
                        '<label class="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Pilih Kelas</label>' +
                        '<select id="nil-kelas" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary outline-none bg-slate-50" onchange="updateNilaiMapelDropdown()"><option value="">-- Pilih --</option>' + db.kelas.map(k => '<option value="' + k.id + '">' + k.nama + '</option>').join('') + '</select>' +
                    '</div>' +
                    '<div>' +
                        '<label class="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Mata Pelajaran <span class="text-[10px] lowercase text-primary font-normal">(Sesuai Jadwal)</span></label>' +
                        '<select id="nil-mapel" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary outline-none bg-slate-50 bg-white" onchange="loadNilaiList()"><option value="">-- Pilih Kelas Dulu --</option></select>' +
                    '</div>' +
                    '<div>' +
                        '<label class="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Tujuan Pembelajaran</label>' +
                        '<input type="text" id="nil-tp" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary outline-none bg-slate-50" placeholder="Contoh: TP 1 Aljabar" onchange="loadNilaiList()"/>' +
                    '</div>' +
                    '<div>' +
                        '<label class="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Jenis Penilaian</label>' +
                        '<select id="nil-jenis" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary outline-none bg-slate-50 mb-2" onchange="toggleNilaiLainnya()">' +
                            '<option value="Tugas">Tugas Harian</option><option value="Ulangan">Ulangan Harian</option><option value="UTS">Penilaian Tengah Semester</option><option value="UAS">Penilaian Akhir Semester</option><option value="Lainnya">Lainnya (Isi Sendiri)</option>' +
                        '</select>' +
                        '<input type="text" id="nil-jenis-custom" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary outline-none bg-white hidden" placeholder="Sebutkan jenis..." onchange="loadNilaiList()"/>' +
                    '</div>' +
                '</div>' +
                '<div id="nil-list-container"></div>' +
            '</div>' +
            '<div id="tab-nilai-rekap" class="tab-content hidden">' +
                '<div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 mb-6 flex flex-col md:flex-row justify-between items-end gap-5 no-print">' +
                    '<div class="grid grid-cols-1 md:grid-cols-2 gap-5 flex-1">' +
                        '<div>' +
                            '<label class="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Pilih Kelas</label>' +
                            '<select id="rek-nil-kelas" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary outline-none bg-slate-50" onchange="updateRekapNilaiMapelDropdown()"><option value="">-- Pilih --</option>' + db.kelas.map(k => '<option value="' + k.id + '">' + k.nama + '</option>').join('') + '</select>' +
                        '</div>' +
                        '<div>' +
                            '<label class="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Mata Pelajaran <span class="text-[10px] lowercase text-primary font-normal">(Sesuai Jadwal)</span></label>' +
                            '<select id="rek-nil-mapel" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary outline-none bg-slate-50 bg-white" onchange="loadRekapNilai()"><option value="">-- Pilih Kelas Dulu --</option></select>' +
                        '</div>' +
                    '</div>' +
                    '<button onclick="cetakRekapNilai()" class="bg-slate-800 text-white font-medium px-6 py-2.5 rounded-xl shadow-md hover:bg-slate-700 transition flex items-center"><i class="fa-solid fa-print mr-2"></i> Cetak</button>' +
                '</div>' +
                '<div id="rek-nil-container"></div>' +
            '</div>';
        }

        function toggleNilaiLainnya() {
            const val = document.getElementById('nil-jenis').value;
            const inputCustom = document.getElementById('nil-jenis-custom');
            if(val === 'Lainnya') { inputCustom.classList.remove('hidden'); } 
            else { inputCustom.classList.add('hidden'); loadNilaiList(); }
        }

        function updateNilaiMapelDropdown() {
            const id_kelas = document.getElementById('nil-kelas').value;
            const selectMapel = document.getElementById('nil-mapel');
            if(!id_kelas) { selectMapel.innerHTML = '<option value="">-- Pilih Kelas Dulu --</option>'; return loadNilaiList(); }
            
            const mapelUnik = [...new Set(db.jadwal.filter(j => j.id_kelas === id_kelas).map(j => j.mapel))];
            
            if(mapelUnik.length === 0) {
                selectMapel.innerHTML = '<option value="">(Belum Ada Jadwal Mapel)</option>';
            } else {
                selectMapel.innerHTML = '<option value="">-- Pilih Mapel --</option>' + mapelUnik.map(m => '<option value="' + m + '">' + m + '</option>').join('');
            }
            loadNilaiList();
        }

        function updateRekapNilaiMapelDropdown() {
            const id_kelas = document.getElementById('rek-nil-kelas').value;
            const selectMapel = document.getElementById('rek-nil-mapel');
            if(!id_kelas) { selectMapel.innerHTML = '<option value="">-- Pilih Kelas Dulu --</option>'; return loadRekapNilai(); }
            
            const mapelUnik = [...new Set(db.jadwal.filter(j => j.id_kelas === id_kelas).map(j => j.mapel))];
            if(mapelUnik.length === 0) selectMapel.innerHTML = '<option value="">(Belum Ada Jadwal Mapel)</option>';
            else selectMapel.innerHTML = '<option value="">-- Semua Mapel Tersedia --</option>' + mapelUnik.map(m => '<option value="' + m + '">' + m + '</option>').join('');
            
            loadRekapNilai();
        }

        function loadNilaiList() {
            const id_kelas = document.getElementById('nil-kelas').value;
            const mapel = document.getElementById('nil-mapel').value;
            const tp = document.getElementById('nil-tp').value;
            let jenis = document.getElementById('nil-jenis').value;
            if(jenis === 'Lainnya') jenis = document.getElementById('nil-jenis-custom').value;
            
            if(!id_kelas || !mapel || !tp || !jenis) { document.getElementById('nil-list-container').innerHTML = ''; return; }

            const siswa = db.siswa.filter(s => s.id_kelas === id_kelas);
            let html = '<div class="bg-white shadow-sm border border-slate-200 rounded-2xl overflow-hidden mb-5"><table class="w-full text-left border-collapse text-sm">' +
                '<thead class="bg-slate-50 border-b border-slate-200 text-slate-500 text-xs uppercase tracking-wider font-semibold"><th class="p-4 pl-6">Nama Siswa</th><th class="p-4 pr-6 w-56 text-right">Skor Pencapaian (0-100)</th></thead><tbody class="divide-y divide-slate-100">';
            
            if(siswa.length === 0) html += '<tr><td colspan="2" class="p-6 text-center text-slate-400 italic">Tidak ada siswa di kelas ini.</td></tr>';

            siswa.forEach(s => {
                const idNilai = ('NIL_' + id_kelas + '_' + mapel + '_' + jenis + '_' + tp + '_' + s.nis).replace(/\s+/g, '');
                const existing = db.nilai.find(n => n.id === idNilai) || {skor: ''};
                html += '<tr class="hover:bg-slate-50 transition" data-nis="' + s.nis + '">' +
                    '<td class="p-4 pl-6 font-medium text-slate-700">' + s.nama + ' <span class="text-slate-400 font-normal text-xs ml-1">(' + s.nis + ')</span></td>' +
                    '<td class="p-4 pr-6"><input type="number" class="nil-skor w-full p-2.5 border border-slate-300 rounded-xl focus:ring-primary focus:border-primary outline-none text-sm text-right font-semibold bg-white" value="' + existing.skor + '" placeholder="0 - 100"/></td>' +
                '</tr>';
            });
            html += '</tbody></table></div>';
            if(siswa.length > 0) html += '<button onclick="saveBulkNilai()" class="bg-primary text-white font-medium px-6 py-3 rounded-xl shadow-md shadow-indigo-200 hover:bg-indigo-700 transition flex items-center justify-center w-full md:w-auto"><i class="fa-solid fa-save mr-2"></i> Simpan Nilai Kelas</button>';
            
            document.getElementById('nil-list-container').innerHTML = html;
        }

        async function saveBulkNilai() {
            const id_kelas = document.getElementById('nil-kelas').value;
            const mapel = document.getElementById('nil-mapel').value;
            const tp = document.getElementById('nil-tp').value;
            let jenis = document.getElementById('nil-jenis').value;
            if(jenis === 'Lainnya') jenis = document.getElementById('nil-jenis-custom').value;
            if(!jenis) return alert("Jenis penilaian wajib diisi.");
            
            let dataArray = [];
            document.querySelectorAll('#nil-list-container tbody tr').forEach(tr => {
                if(!tr.hasAttribute('data-nis')) return;
                const nis = tr.getAttribute('data-nis');
                const skor = tr.querySelector('.nil-skor').value;
                if(skor !== '') {
                    const idNilai = ('NIL_' + id_kelas + '_' + mapel + '_' + jenis + '_' + tp + '_' + nis).replace(/\s+/g, '');
                    dataArray.push({ id: idNilai, nis: nis, mapel: mapel, tp: tp, jenis: jenis, skor: skor });
                }
            });

            if(dataArray.length === 0) return alert("Belum ada skor yang diisi.");
            const res = await fetchAPI('save_bulk', { table: 'nilai', dataArray: dataArray });
            if(res.status === 'success') {
                if(!IS_PREVIEW) { 
                    const fresh = await fetchAPI('login', {username: currentUser.username, password: currentUser.password}); 
                    db = fresh.data; 
                    updateSessionStorage();
                }
                alert("Data Nilai berhasil disimpan!");
            }
        }

        function loadRekapNilai() {
            const id_kelas = document.getElementById('rek-nil-kelas').value;
            const mapel = document.getElementById('rek-nil-mapel').value;
            if(!id_kelas || !mapel) { document.getElementById('rek-nil-container').innerHTML = ''; return; }

            const siswaKelas = db.siswa.filter(s => s.id_kelas === id_kelas);
            const nilaiMapel = db.nilai.filter(n => n.mapel.toLowerCase() === mapel.toLowerCase());
            
            const jenisSet = new Set();
            nilaiMapel.forEach(n => jenisSet.add(n.jenis));
            const jenisCols = Array.from(jenisSet);

            let html = '<div class="bg-white shadow-sm border border-slate-200 rounded-2xl overflow-hidden"><table class="w-full text-left border-collapse text-sm">' +
                '<thead class="bg-slate-50 border-b border-slate-200 text-slate-500 text-xs uppercase tracking-wider font-semibold"><th class="p-4 pl-6">Nama Siswa</th>';
            jenisCols.forEach(j => html += '<th class="p-4 text-center">' + j + '</th>');
            html += '<th class="p-4 pr-6 text-center text-primary bg-indigo-50/50">Rata-Rata</th></thead><tbody class="divide-y divide-slate-100">';

            if(siswaKelas.length === 0) html += '<tr><td colspan="' + (jenisCols.length + 2) + '" class="p-6 text-center text-slate-400 italic">Tidak ada data.</td></tr>';

            siswaKelas.forEach(s => {
                html += '<tr class="hover:bg-slate-50 transition"><td class="p-4 pl-6 font-medium text-slate-700">' + s.nama + '</td>';
                let total = 0, count = 0;
                jenisCols.forEach(j => {
                    const nSiswa = nilaiMapel.filter(n => n.nis === s.nis && n.jenis === j);
                    let avgJenis = '-';
                    if(nSiswa.length > 0) {
                        const sum = nSiswa.reduce((acc, curr) => acc + (parseFloat(curr.skor)||0), 0);
                        avgJenis = (sum / nSiswa.length).toFixed(1);
                        total += parseFloat(avgJenis); count++;
                    }
                    html += '<td class="p-4 text-center text-slate-600">' + avgJenis + '</td>';
                });
                const finalAvg = count > 0 ? (total / count).toFixed(1) : '-';
                const avgColor = finalAvg !== '-' && finalAvg < 70 ? 'text-rose-600' : 'text-emerald-600';
                html += '<td class="p-4 pr-6 text-center font-bold bg-indigo-50/30 ' + avgColor + '">' + finalAvg + '</td></tr>';
            });
            html += '</tbody></table></div>';
            document.getElementById('rek-nil-container').innerHTML = html;
        }

        // ==========================================
        // MODAL & FORM LOGIC
        // ==========================================

        function openModal(table, id = null) {
            currentModalContext = table; currentEditingId = id;
            const modal = document.getElementById('universal-modal');
            document.getElementById('modal-title').innerText = (id ? 'Edit Data ' : 'Tambah Data ') + table.toUpperCase();
            
            let record = id ? (db[table].find(r => r.id === id) || {}) : {};
            let formHtml = '';
            
            if (table === 'kelas') {
                formHtml = createInput('Nama Kelas', 'nama', record.nama);
            } 
            else if (table === 'siswa') {
                formHtml = createInput('NIS', 'nis', record.nis) +
                           createInput('Nama Lengkap Siswa', 'nama', record.nama) +
                           createSelect('Tugaskan ke Kelas', 'id_kelas', db.kelas.map(k => ({val: k.id, label: k.nama})), record.id_kelas) +
                           createCheckbox('Tandai Butuh Bimbingan Wali?', 'is_bimbingan', record.is_bimbingan === "true" || record.is_bimbingan === true);
            }
            else if (table === 'jadwal') {
                formHtml = createSelect('Pilih Kelas', 'id_kelas', db.kelas.map(k => ({val: k.id, label: k.nama})), record.id_kelas) +
                           createInput('Mata Pelajaran', 'mapel', record.mapel, 'text', 'Contoh: Matematika Peminatan') +
                           createSelect('Hari', 'hari', ['Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'].map(h=>({val:h,label:h})), record.hari) +
                           createInput('Jam Mengajar', 'jam', record.jam, 'text', 'Contoh: 08:00 - 09:30');
            }
            else if (table === 'jurnal') {
                const unikMapel = [...new Set(db.jadwal.map(j => j.mapel))];
                const optMapel = unikMapel.length > 0 ? unikMapel.map(m=>({val:m, label:m})) : [{val:'', label:'(Harap Isi Data Jadwal Dulu)'}];

                formHtml = createInput('Tanggal Pelaksanaan', 'tanggal', cleanDate(record.tanggal) || cleanDate(new Date().toISOString()), 'date') +
                           createSelect('Kelas', 'kelas', db.kelas.map(k=>({val: k.id, label: k.nama})), record.kelas) +
                           createSelect('Mata Pelajaran (Dari Jadwal)', 'mapel', optMapel, record.mapel) +
                           createInput('Tujuan Pembelajaran', 'tp', record.tp, 'text', 'Contoh: Memahami Fungsi Kuadrat') +
                           createTextArea('Deskripsi Kegiatan & Hambatan', 'deskripsi', record.deskripsi);
            }
            else if (table === 'bimbingan') {
                const siswaBimbingan = db.siswa.filter(s => s.is_bimbingan === "true" || s.is_bimbingan === true);
                const optSiswa = siswaBimbingan.length > 0 ? siswaBimbingan.map(s => ({val: s.nis, label: s.nama + ' ('+s.nis+')'})) : [{val:'', label:'(Tidak Ada Siswa Butuh Bimbingan)'}];
                
                formHtml = createInput('Tanggal Bimbingan', 'tanggal', cleanDate(record.tanggal) || cleanDate(new Date().toISOString()), 'date') +
                           createSelect('Siswa Terdaftar Bimbingan', 'nis', optSiswa, record.nis) +
                           createSelect('Kategori / Jenis', 'jenis', ['Akademik','Karakter','Keterampilan','Lainnya'].map(j=>({val:j,label:j})), record.jenis) +
                           createTextArea('Catatan Konseling', 'catatan', record.catatan) +
                           createInput('Tindak Lanjut / Solusi', 'tindak_lanjut', record.tindak_lanjut, 'text', 'Contoh: Pemanggilan Orang Tua');
            }
            else if (table === 'dokumen') {
                formHtml = createInput('Nama Dokumen', 'nama', record.nama, 'text', 'Contoh: RPP Matematika') +
                           createInput('Link/URL Dokumen', 'link', record.link, 'url', 'https://...');
            }

            document.getElementById('modal-body').innerHTML = '<form id="dynamic-form" class="space-y-4">' + formHtml + '</form>';
            modal.classList.remove('hidden');
        }

        function createInput(label, name, value = '', type = 'text', placeholder = '') {
            return '<div>' +
                '<label class="block text-xs font-bold text-slate-600 uppercase tracking-wider mb-2">' + label + '</label>' +
                '<input type="' + type + '" name="' + name + '" value="' + value + '" placeholder="' + placeholder + '" class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary focus:border-transparent outline-none transition bg-slate-50 focus:bg-white text-sm" required="required"/>' +
            '</div>';
        }
        function createTextArea(label, name, value = '') {
            return '<div>' +
                '<label class="block text-xs font-bold text-slate-600 uppercase tracking-wider mb-2">' + label + '</label>' +
                '<textarea name="' + name + '" rows="3" class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary focus:border-transparent outline-none transition bg-slate-50 focus:bg-white text-sm" required="required">' + value + '</textarea>' +
            '</div>';
        }
        function createSelect(label, name, options, selectedVal = '') {
            let opts = options.map(o => '<option value="' + o.val + '" ' + (o.val === selectedVal ? 'selected="selected"' : '') + '>' + o.label + '</option>').join('');
            return '<div>' +
                '<label class="block text-xs font-bold text-slate-600 uppercase tracking-wider mb-2">' + label + '</label>' +
                '<select name="' + name + '" class="w-full px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-primary focus:border-transparent outline-none transition bg-slate-50 focus:bg-white text-sm" required="required"><option value="">Pilih Opsi...</option>' + opts + '</select>' +
            '</div>';
        }
        function createCheckbox(label, name, isChecked) {
            return '<div class="flex items-center p-4 bg-slate-50 border border-slate-200 rounded-xl">' +
                '<input type="checkbox" name="' + name + '" value="true" class="w-5 h-5 text-primary border-slate-300 rounded focus:ring-primary cursor-pointer" ' + (isChecked ? 'checked="checked"' : '') + '/>' +
                '<label class="ml-3 text-sm font-semibold text-slate-700 cursor-pointer select-none">' + label + '</label>' +
            '</div>';
        }

        function closeModal() {
            const modal = document.getElementById('universal-modal');
            const box = modal.querySelector('.bg-white');
            box.classList.remove('fade-in');
            modal.classList.add('hidden');
            currentModalContext = ''; currentEditingId = null;
        }

        async function submitModal() {
            const form = document.getElementById('dynamic-form');
            if(!form.checkValidity()) return form.reportValidity();

            const formData = new FormData(form);
            let data = {};
            formData.forEach((value, key) => data[key] = value);
            
            if(currentModalContext === 'siswa' && !data.is_bimbingan) data.is_bimbingan = "false";

            const btn = document.getElementById('modal-submit-btn');
            const originalText = btn.innerHTML;
            btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin mr-2"></i> Menyimpan...'; btn.disabled = true;

            const res = await fetchAPI('save', { table: currentModalContext, id: currentEditingId, data: data });
            btn.innerHTML = originalText; btn.disabled = false;

            if(res.status === 'success') {
                closeModal();
                if(!IS_PREVIEW) { 
                    const fresh = await fetchAPI('login', {username: currentUser.username, password: currentUser.password}); 
                    db = fresh.data; 
                    updateSessionStorage();
                }
                loadMenu(currentMenu);
            } else alert("Gagal menyimpan data ke server.");
        }

        async function deleteData(table, id) {
            if(!confirm("Anda yakin ingin menghapus data ini secara permanen?")) return;
            const res = await fetchAPI('delete', { table: table, id: id });
            if(res.status === 'success') {
                if(!IS_PREVIEW) { 
                    const fresh = await fetchAPI('login', {username: currentUser.username, password: currentUser.password}); 
                    db = fresh.data; 
                    updateSessionStorage();
                }
                loadMenu(currentMenu);
            } else alert("Gagal menghapus data dari server.");
        }

        // ==========================================
        // FITUR TAMBAHAN (PENGATURAN, IMPORT, CETAK)
        // ==========================================

        function renderPengaturan() {
            let set = (db.pengaturan && db.pengaturan.length > 0) ? db.pengaturan[0] : {};
            return '<div class="bg-white p-8 rounded-2xl shadow-sm border border-slate-200 max-w-3xl">' +
                '<h3 class="text-xl font-bold text-slate-800 mb-6 pb-4 border-b border-slate-200"><i class="fa-solid fa-building-columns mr-2 text-primary"></i> Identitas &amp; Kop Surat Sekolah</h3>' +
                '<form id="form-pengaturan">' +
                    '<div class="grid grid-cols-1 gap-6 mb-6">' +
                        createInput('URL/Link Gambar Logo Sekolah (Biarkan kosong untuk default icon)', 'logo', set.logo, 'url', 'https://...') +
                    '</div>' +
                    '<div class="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">' +
                        createInput('Nama Sekolah', 'nama_sekolah', set.nama_sekolah, 'text', 'Contoh: SMA Negeri 1') +
                        createInput('Alamat Lengkap', 'alamat', set.alamat, 'text', 'Contoh: Jl. Pendidikan No. 1') +
                        createInput('Nama Kepala Sekolah', 'kepsek', set.kepsek, 'text') +
                        createInput('NIP Kepala Sekolah', 'nip_kepsek', set.nip_kepsek, 'text') +
                        createInput('Nama Guru / Wali Kelas', 'guru', set.guru, 'text') +
                        createInput('NIP Guru', 'nip_guru', set.nip_guru, 'text') +
                    '</div>' +
                    '<div class="mt-8 flex justify-end">' +
                        '<button type="button" onclick="savePengaturan()" class="bg-primary text-white font-medium px-8 py-3 rounded-xl shadow-md hover:bg-indigo-700 transition" id="btn-save-pengaturan"><i class="fa-solid fa-save mr-2"></i> Simpan Pengaturan</button>' +
                    '</div>' +
                '</form>' +
            '</div>';
        }

        async function savePengaturan() {
            const form = document.getElementById('form-pengaturan');
            if(!form.checkValidity()) return form.reportValidity();
            
            let data = {};
            new FormData(form).forEach((value, key) => data[key] = value);
            
            const btn = document.getElementById('btn-save-pengaturan');
            btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin mr-2"></i> Menyimpan...'; btn.disabled = true;

            const existingId = (db.pengaturan && db.pengaturan.length > 0) ? db.pengaturan[0].id : null;
            const res = await fetchAPI('save', { table: 'pengaturan', id: existingId, data: data });
            
            if(res.status === 'success') {
                if(!IS_PREVIEW) { 
                    const fresh = await fetchAPI('login', {username: currentUser.username, password: currentUser.password}); 
                    db = fresh.data; 
                    updateSessionStorage();
                }
                else { 
                    if(existingId) db.pengaturan[0] = {...db.pengaturan[0], ...data}; 
                    else db.pengaturan.push({id: "P1", ...data}); 
                }
                
                applyPengaturanUI();
                alert("Pengaturan sekolah berhasil disimpan.");
            }
            btn.innerHTML = '<i class="fa-solid fa-save mr-2"></i> Simpan Pengaturan'; btn.disabled = false;
        }

        function downloadFormatExcel() {
            const ws_data = [
                ["NIS", "Nama Siswa", "Kelas", "Bimbingan?"],
                ["1001", "Siswa Contoh Satu", "10 IPA 1", "Tidak"],
                ["1002", "Siswa Contoh Dua", "10 IPA 2", "Ya"],
                ["1003", "Siswa Contoh Tiga", "10 IPA 1", "Tidak"]
            ];
            const ws = XLSX.utils.aoa_to_sheet(ws_data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "FormatSiswa");
            XLSX.writeFile(wb, "Format_Import_Siswa.xlsx");
        }

        function triggerImportExcel() { document.getElementById('excel-file').click(); }

        function handleImportExcel(event) {
            const file = event.target.files[0];
            if(!file) return;
            const reader = new FileReader();
            reader.onload = async function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, {type: 'array'});
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const rows = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                    
                    if(rows.length <= 1) {
                        alert("File Excel kosong atau format salah.");
                        return;
                    }

                    let dataArray = [];
                    let missingClasses = new Set();

                    for(let i=1; i<rows.length; i++) {
                        const row = rows[i];
                        if(!row || row.length === 0 || (!row[0] && !row[1])) continue;

                        const nis = String(row[0] || "").trim();
                        const nama = String(row[1] || "").trim();
                        const namaKelas = String(row[2] || "").trim();
                        const bimbingan = String(row[3] || "").trim().toLowerCase();
                        
                        const kelasObj = db.kelas.find(k => k.nama.toLowerCase() === namaKelas.toLowerCase());
                        let id_kelas = "";
                        if(kelasObj) {
                            id_kelas = kelasObj.id;
                        } else {
                            missingClasses.add(namaKelas);
                            continue;
                        }

                        const is_bimbingan = (bimbingan === 'ya' || bimbingan === 'true' || bimbingan === '1') ? "true" : "false";

                        if(nis && nama && id_kelas) {
                            dataArray.push({
                                id: "ID_" + new Date().getTime() + "_" + i,
                                nis: nis,
                                nama: nama,
                                id_kelas: id_kelas,
                                is_bimbingan: is_bimbingan
                            });
                        }
                    }

                    if(missingClasses.size > 0) {
                        alert("Peringatan: Beberapa data dilewati karena Kelas tidak ditemukan di sistem:\n" + Array.from(missingClasses).join(', '));
                    }

                    if(dataArray.length > 0) {
                        showLoading(true);
                        const res = await fetchAPI('save_bulk', { table: 'siswa', dataArray: dataArray });
                        showLoading(false);
                        if(res.status === 'success') {
                            if(!IS_PREVIEW) { 
                                const fresh = await fetchAPI('login', {username: currentUser.username, password: currentUser.password}); 
                                db = fresh.data;
                                updateSessionStorage();
                            }
                            alert(dataArray.length + " Data siswa berhasil diimpor!");
                            loadMenu('siswa');
                        }
                    } else {
                        alert("Gagal impor: Pastikan kolom Kelas pada file Excel sama persis dengan nama kelas di Master Data Kelas.");
                    }
                } catch (error) {
                    alert("Gagal memproses file Excel. Pastikan file berformat .xlsx");
                    console.error(error);
                }
                document.getElementById('excel-file').value = "";
            };
            reader.readAsArrayBuffer(file);
        }

        function cetakRekapAbsen() {
            const klsDropdown = document.getElementById('rek-abs-kelas');
            const namaKelas = klsDropdown.options[klsDropdown.selectedIndex].text;
            const tipe = document.getElementById('rek-abs-tipe').value;
            let periodeStr = '-';
            
            if (tipe === 'semester') {
                const t = document.getElementById('rek-abs-tahun')?.value;
                const s = document.getElementById('rek-abs-smt')?.value;
                if(t && s) periodeStr = 'Semester ' + (s === 'ganjil' ? 'Ganjil' : 'Genap') + ' ' + t;
            } else {
                periodeStr = document.getElementById('rek-abs-waktu')?.value || '-';
            }

            cetakLaporan('rek-abs-container', 'Rekapitulasi Kehadiran Siswa', {
                'Kelas / Rombel': namaKelas.includes('--') ? 'Semua Kelas' : namaKelas,
                'Waktu / Periode': periodeStr
            });
        }

        function cetakRekapNilai() {
            const klsDropdown = document.getElementById('rek-nil-kelas');
            const namaKelas = klsDropdown.options[klsDropdown.selectedIndex].text;
            const mapel = document.getElementById('rek-nil-mapel').value;
            cetakLaporan('rek-nil-container', 'Rekapitulasi Penilaian Siswa', {
                'Kelas / Rombel': namaKelas.includes('--') ? 'Belum Dipilih' : namaKelas,
                'Mata Pelajaran': mapel || '-'
            });
        }

        function cetakLaporan(elementId, title, extraInfo = {}) {
            let content = document.getElementById(elementId).innerHTML;
            content = content.replace(/<button[^>]*>[\s\S]*?<\/button>/gi, ''); 
            content = content.replace(/<td[^>]*no-print[^>]*>[\s\S]*?<\/td>/gi, ''); 
            content = content.replace(/<th[^>]*>Aksi<\/th>/gi, ''); 
            
            let set = (db.pengaturan && db.pengaturan.length > 0) ? db.pengaturan[0] : {nama_sekolah: 'Nama Sekolah', alamat: 'Alamat Sekolah', kepsek: 'Kepala Sekolah', nip_kepsek: '-', guru: currentUser.nama, nip_guru: '-'};
            
            let logoHtml = set.logo ? '<img src="' + set.logo + '" style="width: 80px; height: 80px; object-fit: contain; position: absolute; left: 0; top: 0;"/>' : '';

            let extraHtml = '';
            if (Object.keys(extraInfo).length > 0) {
                extraHtml = '<table style="width:50%; border:none; margin-bottom:15px; font-size:14px; text-align:left;">';
                for (const [key, value] of Object.entries(extraInfo)) {
                    extraHtml += '<tr><td style="border:none; padding:3px 0; font-weight:bold; width:150px;">' + key + '</td><td style="border:none; padding:3px 0;">: ' + value + '</td></tr>';
                }
                extraHtml += '</table>';
            }

            let win = window.open('', '', 'width=900,height=700');
            win.document.write(`
                <html><head><title>Cetak - ` + title + `</title>
                <style>
                    body { font-family: 'Times New Roman', Times, serif; padding: 20px; color: #000; }
                    .kop { text-align: center; border-bottom: 4px solid #000; padding-bottom: 10px; margin-bottom: 20px; position: relative; min-height: 80px; }
                    .kop h2 { margin: 0; font-size: 24px; text-transform: uppercase; letter-spacing: 1px; }
                    .kop p { margin: 5px 0 0 0; font-size: 14px; }
                    h3 { text-align: center; margin-bottom: 25px; font-size: 18px; text-transform: uppercase; }
                    table { width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 30px; }
                    th, td { border: 1px solid #000; padding: 8px 12px; text-align: left; }
                    th { font-weight: bold; text-align: center; background-color: #f8fafc !important; -webkit-print-color-adjust: exact; color-adjust: exact; }
                    .ttd-container { margin-top: 50px; width: 100%; display: flex; justify-content: space-between; font-size: 14px; page-break-inside: avoid; }
                    .ttd-box { text-align: center; width: 250px; }
                    .ttd-nama { font-weight: bold; text-decoration: underline; margin-top: 70px; }
                </style>
                </head><body>
                    <div class="kop">
                        ` + logoHtml + `
                        <h2>` + set.nama_sekolah + `</h2>
                        <p>` + set.alamat + `</p>
                    </div>
                    <h3>` + title + `</h3>
                    ` + extraHtml + `
                    ` + content + `
                    <div class="ttd-container">
                        <div class="ttd-box">
                            <p>Mengetahui,<br/>Kepala Sekolah</p>
                            <div class="ttd-nama">` + (set.kepsek || '.......................') + `</div>
                            <div>NIP. ` + (set.nip_kepsek || '-') + `</div>
                        </div>
                        <div class="ttd-box">
                            <p>Mengesahkan,<br/>Guru Mata Pelajaran</p>
                            <div class="ttd-nama">` + (set.guru || currentUser.nama) + `</div>
                            <div>NIP. ` + (set.nip_guru || '-') + `</div>
                        </div>
                    </div>
                    <script>setTimeout(()=> { window.print(); window.close(); }, 800);<\/script>
                </body></html>
            `);
            win.document.close();
        }
    //]]>
    </script>
</body>
</html>

