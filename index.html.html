<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>REB/NESA CBC Lesson Plan Generator</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:"Times New Roman",Times,serif;background:linear-gradient(135deg,#1e3a8a,#5b21b6);padding:16px;min-height:100vh}
.container{max-width:1100px;margin:0 auto}
.header{background:#fff;border-radius:12px;padding:20px;margin-bottom:14px;text-align:center;box-shadow:0 6px 20px rgba(0,0,0,.15)}
.header h1{color:#1e3a8a;font-size:1.7rem}
.header p{color:#555;margin-top:4px}
.card{background:#fff;border-radius:12px;padding:20px 22px;margin-bottom:14px;box-shadow:0 6px 20px rgba(0,0,0,.12)}
.lang-box{background:#eff6ff;border:2px solid #3b82f6;border-radius:8px;padding:10px 12px;margin-bottom:14px}
.form-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:12px}
.form-group{display:flex;flex-direction:column}
.form-group.full{grid-column:1/-1}
label{font-weight:700;font-size:.86rem;margin-bottom:3px;color:#1e293b}
label.req::after{content:" *";color:#dc2626}
label.opt{color:#2563eb}
input,select,textarea{padding:8px 10px;border:2px solid #d1d5db;border-radius:7px;font-size:.93rem;font-family:inherit}
input:focus,select:focus,textarea:focus{outline:none;border-color:#3b82f6}
.double{display:flex;gap:8px}.double input{flex:1}
.btn-row{display:flex;flex-wrap:wrap;gap:8px;justify-content:center;margin-top:16px}
.btn{border:none;border-radius:8px;padding:11px 18px;font-weight:700;cursor:pointer;font-size:.95rem}
.btn-gen{background:linear-gradient(135deg,#1e40af,#7c3aed);color:#fff}
.btn-sample{background:#e0e7ff;color:#1e3a8a}
.btn-clear{background:#f1f5f9;color:#334155}
.btn-pdf{background:#dc2626;color:#fff}
.btn-word{background:#059669;color:#fff}
.btn-excel{background:#16a34a;color:#fff}
.btn-csv{background:#0d9488;color:#fff}
.btn-print{background:#0f172a;color:#fff}
.error{color:#dc2626;text-align:center;margin-top:8px;min-height:1.2em}
.result{display:none}.result.show{display:block}
.toolbar{display:flex;flex-wrap:wrap;gap:8px;justify-content:center;margin-bottom:12px}
.dl-choice{background:#f8fafc;border:1px solid #cbd5e1;border-radius:8px;padding:10px;margin-bottom:10px;text-align:center}
.dl-choice label{font-weight:600;margin-right:12px;font-size:.9rem}
.lesson-plan{font-size:12.5pt;line-height:1.38;color:#000;background:#fff}
.lesson-plan .title{text-align:center;font-weight:700;font-size:17pt;margin-bottom:10px;letter-spacing:.03em}
.lesson-plan .meta{margin-bottom:7px;font-size:13pt}
.lesson-plan .meta b{font-weight:700}
.lesson-plan table{width:100%;border-collapse:collapse;margin-bottom:0;table-layout:fixed}
.lesson-plan td,.lesson-plan th{border:1.4px solid #222;padding:6px 7px;vertical-align:top;font-size:12pt}
.lesson-plan .hdr{background:#d0d5dd;font-weight:700;font-size:12pt}
.lesson-plan .lab{background:#e8eaed;font-weight:700;width:32%;font-size:12pt}
.lesson-plan ul{margin:2px 0 2px 4px;padding:0;list-style:none}
.lesson-plan li{margin:2px 0;font-size:12pt;padding-left:0}
.lesson-plan li::before{content:none}
.lesson-plan .plan-footer{text-align:right;margin-top:2px;padding-top:3px;border-top:1px solid #888;font-size:8.5pt;font-weight:600;color:#475569;letter-spacing:.02em;overflow:auto}
.lesson-plan .plan-footer::after{content:"";display:table;clear:both}
.hint{font-size:.78rem;color:#64748b;margin-top:2px}
.lesson-plan.editing{outline:2px dashed #059669;outline-offset:4px;background:#f0fdf4}
.lesson-plan.editing td,.lesson-plan.editing .meta,.lesson-plan.editing .title{cursor:text}

@media print{
  body{background:#fff;padding:0;margin:0}
  .no-print{display:none!important}
  .card,.result{box-shadow:none;border-radius:0;padding:0;margin:0}
  .container{max-width:none;margin:0}
  .lesson-plan{font-size:10pt;line-height:1.22}
  .lesson-plan td,.lesson-plan th{font-size:9.2pt;padding:3.5px 4px}
  .lesson-plan .title{font-size:12.5pt;margin-bottom:4px}
  .lesson-plan .meta{font-size:10pt}
  .lesson-plan .plan-footer{font-size:8pt;margin-top:1px;padding-top:2px;text-align:right}
  .lesson-plan ul{margin:1px 0 1px 12px}
  .lesson-plan li{margin:1px 0}
  @page{size:A4 portrait;margin:7mm}
}

/* ===== AUTH / LOGIN / ADMIN ===== */
#authScreen,#adminScreen{display:none;max-width:480px;margin:40px auto}
#authScreen.active,#adminScreen.active{display:block}
#appScreen{display:none}
#appScreen.active{display:block}
.auth-card{background:#fff;border-radius:14px;padding:28px 26px;box-shadow:0 10px 40px rgba(0,0,0,.25)}
.auth-card h2{color:#1e3a8a;text-align:center;margin-bottom:6px;font-size:1.5rem}
.auth-card .sub{text-align:center;color:#64748b;font-size:.9rem;margin-bottom:18px}
.auth-tabs{display:flex;gap:0;margin-bottom:18px;border-radius:8px;overflow:hidden;border:2px solid #1e3a8a}
.auth-tabs button{flex:1;padding:10px;border:none;background:#eff6ff;color:#1e3a8a;font-weight:700;cursor:pointer;font-size:.95rem}
.auth-tabs button.active{background:#1e3a8a;color:#fff}
.auth-form .form-group{margin-bottom:12px}
.auth-form label{font-weight:700;font-size:.86rem;display:block;margin-bottom:3px;color:#1e293b}
.auth-form input{width:100%;padding:10px 12px;border:1.5px solid #cbd5e1;border-radius:8px;font-size:.95rem}
.auth-form input:focus{outline:none;border-color:#3b82f6;box-shadow:0 0 0 3px rgba(59,130,246,.2)}
.auth-btn{width:100%;padding:12px;border:none;border-radius:8px;background:#1e3a8a;color:#fff;font-weight:700;font-size:1rem;cursor:pointer;margin-top:8px}
.auth-btn:hover{background:#1e40af}
.auth-btn.secondary{background:#64748b}
.auth-btn.danger{background:#dc2626}
.auth-btn.success{background:#16a34a}
.auth-msg{padding:10px 12px;border-radius:8px;margin-bottom:12px;font-size:.9rem;display:none}
.auth-msg.error{display:block;background:#fef2f2;color:#b91c1c;border:1px solid #fecaca}
.auth-msg.ok{display:block;background:#f0fdf4;color:#166534;border:1px solid #bbf7d0}
.auth-msg.info{display:block;background:#eff6ff;color:#1e40af;border:1px solid #bfdbfe}
.user-bar{display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap;background:#1e3a8a;color:#fff;padding:10px 16px;border-radius:10px;margin-bottom:14px}
.user-bar .left{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.user-bar .badge{background:#fbbf24;color:#1e3a8a;font-weight:700;font-size:.75rem;padding:2px 8px;border-radius:999px}
.user-bar .badge.teacher{background:#93c5fd;color:#1e3a8a}
.user-bar button{padding:6px 12px;border:none;border-radius:6px;font-weight:700;cursor:pointer;font-size:.85rem}
.user-bar .btn-out{background:#fee2e2;color:#b91c1c}
.user-bar .btn-admin{background:#fbbf24;color:#1e3a8a}
.admin-table{width:100%;border-collapse:collapse;font-size:.88rem;margin-top:12px}
.admin-table th,.admin-table td{border:1px solid #e2e8f0;padding:8px 10px;text-align:left}
.admin-table th{background:#1e3a8a;color:#fff}
.admin-table tr:nth-child(even){background:#f8fafc}
.admin-table .status{font-weight:700;font-size:.8rem;padding:2px 8px;border-radius:999px}
.admin-table .status.pending{background:#fef3c7;color:#92400e}
.admin-table .status.approved{background:#dcfce7;color:#166534}
.admin-table .status.rejected{background:#fee2e2;color:#b91c1c}
.admin-table .status.disabled{background:#e2e8f0;color:#475569}
.admin-actions{display:flex;gap:6px;flex-wrap:wrap}
.admin-actions button{padding:4px 10px;border:none;border-radius:5px;font-size:.8rem;font-weight:700;cursor:pointer}
.admin-actions .approve{background:#16a34a;color:#fff}
.admin-actions .reject{background:#dc2626;color:#fff}
.admin-actions .disable{background:#64748b;color:#fff}
.admin-actions .enable{background:#2563eb;color:#fff}
.admin-actions .delete{background:#7f1d1d;color:#fff}
.admin-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;flex-wrap:wrap;gap:8px}
.hint-box{background:#eff6ff;border:1px solid #93c5fd;border-radius:8px;padding:10px 12px;font-size:.85rem;color:#1e40af;margin-bottom:14px}


/* Password modal */
.pw-modal-bg{display:none;position:fixed;inset:0;background:rgba(15,23,42,.55);z-index:9999;align-items:center;justify-content:center;padding:16px}
.pw-modal-bg.open{display:flex}
.pw-modal{background:#fff;border-radius:12px;padding:22px 20px;max-width:420px;width:100%;box-shadow:0 20px 50px rgba(0,0,0,.3)}
.pw-modal h3{color:#1e3a8a;margin:0 0 10px;font-size:1.15rem}
.pw-modal .pw-row{display:flex;gap:8px;align-items:center;margin:10px 0}
.pw-modal input[type=text],.pw-modal input[type=password]{flex:1;padding:10px;border:1.5px solid #cbd5e1;border-radius:8px;font-size:.95rem}
.pw-modal .pw-actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:14px}
.pw-modal .pw-actions button{padding:8px 14px;border:none;border-radius:7px;font-weight:700;cursor:pointer;font-size:.88rem}
.pw-show{background:#eff6ff;border:1px solid #93c5fd;border-radius:8px;padding:10px 12px;font-family:monospace;font-size:1rem;word-break:break-all;margin:8px 0}
.login-links{text-align:center;margin-top:14px;font-size:.88rem}
.login-links a{color:#1e3a8a;font-weight:700;cursor:pointer;text-decoration:underline}


#lessonPlanContent{font-size:15px;line-height:1.35}
#lessonPlanContent table{font-size:14px}
#lessonPlanContent .title{font-size:1.35rem}
#lessonPlanContent .meta{font-size:14px}
#lessonPlanContent td{padding:5px 7px}

</style>
</head>
<body>

<!-- ===== AUTH SCREEN ===== -->

<!-- Password view / reset modal -->
<div class="pw-modal-bg" id="pwModalBg" onclick="if(event.target===this)closePwModal()">
  <div class="pw-modal" role="dialog" aria-modal="true">
    <h3 id="pwModalTitle">Password</h3>
    <p id="pwModalSub" style="color:#64748b;font-size:.88rem;margin:0 0 8px"></p>
    <div id="pwViewBox" style="display:none">
      <label style="font-weight:700;font-size:.85rem">Current password</label>
      <div class="pw-show" id="pwViewValue">—</div>
      <button type="button" class="auth-btn secondary" style="width:auto;padding:8px 12px" onclick="copyPwView()">Copy password</button>
    </div>
    <div id="pwResetBox" style="display:none">
      <div class="form-group" style="margin-bottom:8px">
        <label style="font-weight:700;font-size:.85rem">New password</label>
        <input type="text" id="pwNewInput" placeholder="Enter new password (min 6)">
      </div>
      <div class="form-group">
        <label style="font-weight:700;font-size:.85rem">Confirm new password</label>
        <input type="text" id="pwNewInput2" placeholder="Confirm">
      </div>
    </div>
    <div id="pwRecoveryBox" style="display:none">
      <div class="hint-box">Enter the <strong>Admin Recovery Code</strong> shown when you first registered. Then set a new admin password.</div>
      <div class="form-group" style="margin-bottom:8px">
        <label style="font-weight:700;font-size:.85rem">Recovery Code</label>
        <input type="text" id="pwRecoveryCode" placeholder="e.g. ADMIN-XXXX-XXXX">
      </div>
      <div class="form-group" style="margin-bottom:8px">
        <label style="font-weight:700;font-size:.85rem">New Admin Password</label>
        <input type="text" id="pwAdminNew" placeholder="Min 6 characters">
      </div>
      <div class="form-group">
        <label style="font-weight:700;font-size:.85rem">Confirm Password</label>
        <input type="text" id="pwAdminNew2" placeholder="Confirm">
      </div>
    </div>
    <div class="pw-actions">
      <button type="button" id="pwBtnSave" class="auth-btn success" style="width:auto;display:none" onclick="savePwModal()">Save</button>
      <button type="button" class="auth-btn secondary" style="width:auto" onclick="closePwModal()">Close</button>
    </div>
  </div>
</div>

<div id="authScreen" class="active">
  <div class="auth-card">
    <h2>REB / NESA CBC Lesson Plan</h2>
    <p class="sub" id="authSub">Sign in to continue</p>
    <div class="auth-tabs">
      <button type="button" id="tabLogin" class="active" onclick="showAuthTab('login')">Login</button>
      <button type="button" id="tabRegister" onclick="showAuthTab('register')">Register</button>
    </div>
    <div id="authMsg" class="auth-msg"></div>

    <form id="loginForm" class="auth-form" onsubmit="return handleLogin(event)">
      <div class="form-group">
        <label>Email / Username</label>
        <input type="text" id="loginEmail" required autocomplete="username" placeholder="you@school.rw">
      </div>
      <div class="form-group">
        <label>Password</label>
        <input type="password" id="loginPassword" required autocomplete="current-password" placeholder="••••••••">
      </div>
      <button type="submit" class="auth-btn">Login</button>
      <div class="login-links">
        <a onclick="openAdminRecovery()">Forgot admin password? Reset with recovery code</a>
      </div>
    </form>

    <form id="registerForm" class="auth-form" style="display:none" onsubmit="return handleRegister(event)">
      <div class="form-group">
        <label>Full Name</label>
        <input type="text" id="regName" required placeholder="e.g. SIBONIYO LAMBERT">
      </div>
      <div class="form-group">
        <label>Email / Username</label>
        <input type="text" id="regEmail" required autocomplete="username" placeholder="you@school.rw">
      </div>
      <div class="form-group">
        <label>Password (min 6 characters)</label>
        <input type="password" id="regPassword" required minlength="6" autocomplete="new-password" placeholder="••••••••">
      </div>
      <div class="form-group">
        <label>Confirm Password</label>
        <input type="password" id="regPassword2" required minlength="6" autocomplete="new-password" placeholder="••••••••">
      </div>
      <div id="regHint" class="hint-box" style="display:none"></div>
      <button type="submit" class="auth-btn" id="regSubmitBtn">Register as Administrator</button>
      <p id="regClosedMsg" style="display:none;margin-top:12px;padding:12px;background:#fef2f2;border:1px solid #fecaca;border-radius:8px;color:#b91c1c;font-size:.9rem;text-align:center">
        <strong>Registration is closed.</strong><br>
        Only the Administrator can create accounts for teachers.<br>
        Contact your Admin to get access.
      </p>
    </form>
  </div>
</div>

<!-- ===== ADMIN SCREEN ===== -->
<div id="adminScreen">
  <div class="auth-card" style="max-width:720px;margin:0 auto">
    <div class="admin-header">
      <h2 style="margin:0;color:#1e3a8a">Admin · User Management</h2>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button type="button" class="auth-btn secondary" style="width:auto;padding:8px 14px;background:#1e40af" onclick="showMyRecoveryCode()">My recovery code</button>
        <button type="button" class="auth-btn secondary" style="width:auto;padding:8px 14px" onclick="closeAdmin()">← Back to Lesson Plan</button>
      </div>
    </div>
    <div class="hint-box">
      <strong>You are the only Administrator.</strong> No one else can register themselves.
      Create accounts for teachers below. Only you can add new users.
    </div>
    <div id="adminMsg" class="auth-msg"></div>

    <div class="card" style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:14px 16px;margin-bottom:16px">
      <h3 style="color:#1e3a8a;margin:0 0 10px;font-size:1.05rem">Register a new teacher</h3>
      <form id="adminCreateForm" onsubmit="return adminCreateUser(event)">
        <div class="form-grid" style="grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px">
          <div class="form-group">
            <label class="req">Full Name</label>
            <input type="text" id="adminNewName" required placeholder="Teacher full name">
          </div>
          <div class="form-group">
            <label class="req">Email / Username</label>
            <input type="text" id="adminNewEmail" required placeholder="teacher@school.rw">
          </div>
          <div class="form-group">
            <label class="req">Password (min 6)</label>
            <input type="password" id="adminNewPass" required minlength="6" placeholder="••••••••">
          </div>
          <div class="form-group">
            <label class="req">Confirm Password</label>
            <input type="password" id="adminNewPass2" required minlength="6" placeholder="••••••••">
          </div>
        </div>
        <button type="submit" class="auth-btn success" style="width:auto;margin-top:10px;padding:10px 18px">+ Create Teacher Account</button>
      </form>
    </div>

    <h3 style="color:#1e3a8a;margin:0 0 8px;font-size:1.05rem">All users</h3>
    <div style="overflow-x:auto">
      <table class="admin-table" id="adminUserTable">
        <thead>
          <tr>
            <th>Name</th>
            <th>Email</th>
            <th>Role</th>
            <th>Status</th>
            <th>Registered</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody id="adminUserBody"></tbody>
      </table>
    </div>
    <p style="margin-top:12px;font-size:.85rem;color:#64748b">Pending users cannot log in until you approve them.</p>
  </div>
</div>


<div id="appScreen">
<div class="container">
  
  <div class="user-bar no-print" id="userBar">
    <div class="left">
      <span id="userBarName">—</span>
      <span class="badge" id="userBarRole">—</span>
    </div>
    <div>
      <button type="button" class="btn-admin" id="btnAdminPanel" style="display:none" onclick="openAdmin()">Admin Panel</button>
      <button type="button" class="btn-out" onclick="handleLogout()">Logout</button>
    </div>
  </div>

  <div class="header no-print">
    <h1>REB / NESA CBC Lesson Plan</h1>
    <p>Complete professional lesson plan · PDF · Word · Excel · CSV</p>
  </div>

  <div class="card no-print">
    <div class="lang-box">
      <div class="form-group">
        <label>Language / Ururimi / Langue / Lugha</label>
        <select id="languageSelect" onchange="onLanguageChange()">
          <option value="en">English</option>
          <option value="rw">Kinyarwanda</option>
          <option value="fr">Français</option>
          <option value="sw">Kiswahili</option>
        </select>
        <span class="hint" style="display:block;margin-top:4px">Changing language updates the whole lesson plan (labels + activities + competences) into that language.</span>
      </div>
    </div>

    <form id="lessonForm" onsubmit="generateLessonPlan(); return false;">
      <div class="form-grid">
        <div class="form-group"><label class="req">School Name</label><input id="schoolName" required placeholder="e.g. BWENDA Primary School"></div>
        <div class="form-group"><label class="req">Teacher's Name</label><input id="teacherName" required placeholder="e.g. SIBONIYO LAMBERT"></div>
        <div class="form-group"><label class="req">Term</label>
          <select id="term" required><option value="">Select</option><option value="1">1</option><option value="2">2</option><option value="3">3</option><option value="I">I</option><option value="II">II</option><option value="III">III</option></select>
        </div>
        <div class="form-group"><label class="req">Date</label><input type="date" id="date" required></div>
                <div class="form-group"><label class="req">Level</label>
          <select id="level" required onchange="onLevelChange()">
            <option value="">Select level</option>
            <option value="Primary">Primary</option>
            <option value="Ordinary Level">Ordinary Level (O-Level)</option>
            <option value="Advanced Level">Advanced Level (A-Level)</option>
          </select>
        </div>
        <div class="form-group"><label class="req">Class</label>
          <select id="class" required onchange="onClassChange()">
            <option value="">Select class</option>
          </select>
        </div>
        <div class="form-group"><label class="req">Class group</label>
          <select id="classGroup" required>
            <option value="">Select group</option>
            <option value="A">A</option>
            <option value="B">B</option>
            <option value="C">C</option>
            <option value="D">D</option>
            <option value="E">E</option>
            <option value="F">F</option>
            <option value="G">G</option>
          </select>
        </div>
        <div class="form-group"><label class="req">Subject</label>
          <select id="subject" required onchange="onSubjectChange()">
            <option value="">Select subject</option>
          </select>
        </div>
        <div class="form-group"><label class="opt">Or type subject</label>
          <input id="subjectCustom" placeholder="Custom subject if not in list" oninput="if(this.value){document.getElementById('subject').value='';}">
        </div>
        <div class="form-group"><label class="req">Unit No</label><input type="number" id="unitNo" min="1" required></div>
        <div class="form-group"><label class="req">Lesson No / Total</label>
          <div class="double"><input type="number" id="lessonNo" min="1" placeholder="No" required><input type="number" id="totalLessons" min="1" placeholder="Total" required></div>
        </div>
        <div class="form-group"><label class="req">Duration (min)</label><input type="number" id="duration" min="20" max="120" required value="40"></div>
        <div class="form-group"><label class="req">Class size</label><input type="number" id="classSize" min="1" required></div>
        <div class="form-group full"><label class="req">Unit title (from REB books)</label>
          <select id="unitTitleSelect" onchange="onUnitSelect()">
            <option value="">Select unit title from REB syllabus (or type below)</option>
          </select>
          <input id="unitTitle" required placeholder="Unit title — select from REB list above or type custom" style="margin-top:6px">
          <span class="hint">REB 2026–2027 content distribution: units & lessons filter by <strong>class + subject</strong> (elearning.reb.rw). Shown as Unit 1: title. Unit No auto-fills.</span>
        </div>
        <div class="form-group full"><label class="req">Title of the lesson (from REB books)</label>
          <select id="lessonTitleSelect" onchange="onLessonSelect()">
            <option value="">Select lesson title from unit (or type below)</option>
          </select>
          <input id="lessonTitle" required placeholder="Lesson title — select from list above or type custom" style="margin-top:6px">
          <span class="hint">Lessons update from the selected unit (Lesson 1, Lesson 2…). Lesson No &amp; total can auto-fill.</span>
        </div>
<div class="form-group full"><label class="opt">Key unit competence (from REB list when you select unit — shown on plan as written; auto only if empty)</label><textarea id="keyCompetence" rows="2" placeholder="Filled from REB when you select a unit — kept as written on the plan. Leave empty only if you want auto-generate."></textarea></div>
        <div class="form-group full"><label class="opt">Instructional objectives (optional — auto if empty)</label><textarea id="objective" rows="2" placeholder="Leave empty — auto: By using [materials], learners will be able to…"></textarea></div>
        <div class="form-group full"><label class="opt">Special educational needs (select)</label>
          <select id="specialNeeds" onchange="onSpecialNeedsChange()">
            <option value="None">None</option>
            <option value="Visual impairment">Visual impairment</option>
            <option value="Hearing impairment">Hearing impairment</option>
            <option value="Physical disability">Physical disability</option>
            <option value="Intellectual disability">Intellectual disability</option>
            <option value="Speech and language difficulties">Speech and language difficulties</option>
            <option value="Learning difficulties (dyslexia / slow learners)">Learning difficulties (dyslexia / slow learners)</option>
            <option value="Autism spectrum">Autism spectrum</option>
            <option value="ADHD / attention difficulties">ADHD / attention difficulties</option>
            <option value="Multiple disabilities">Multiple disabilities</option>
            <option value="Gifted and talented">Gifted and talented</option>
            <option value="Other (specify below)">Other (specify below)</option>
          </select>
          <input id="specialNeedsDetail" placeholder="Number of learners in each category / extra detail (optional)" style="margin-top:6px">
          <span class="hint">Select the category; add numbers e.g. "2 learners with visual impairment"</span>
        </div>
        <div class="form-group full"><label class="opt">Plan of this class (Location)</label>
          <select id="location">
            <option value="Classroom">Classroom (indoors)</option>
            <option value="Outside the classroom">Outside the classroom</option>
            <option value="School garden">School garden</option>
            <option value="School laboratory">School laboratory</option>
            <option value="Computer lab / ICT room">Computer lab / ICT room</option>
            <option value="Library">Library</option>
            <option value="Playground / sports field">Playground / sports field</option>
            <option value="Workshop">Workshop</option>
            <option value="Classroom and outside">Classroom and outside (mixed)</option>
            <option value="Other">Other</option>
          </select>
          <span class="hint">Where the lesson will take place (in / outside)</span>
        </div>
        <div class="form-group full"><label class="opt">Learning materials</label><input id="materials" placeholder="Auto from subject + lesson if empty"></div>
        <div class="form-group full"><label class="opt">References</label><textarea id="references" rows="2" placeholder="Auto REB-style references if empty"></textarea></div>
      </div>
      <div class="btn-row">
        <button type="button" class="btn btn-sample" onclick="fillSample()">Load sample (like PDF)</button>
        <button type="button" class="btn btn-clear" onclick="clearForm()">Clear</button>
        <button type="button" class="btn btn-gen" onclick="generateLessonPlan()">⚡ Generate with STRONG POWER (fill &amp; complete)</button>
      </div>
      <div id="errorMessage" class="error"></div>
    </form>
  </div>

  <div id="resultContainer" class="card result">
    <div class="no-print" id="editPlanBox" style="background:#ecfdf5;border:2px solid #059669;border-radius:8px;padding:12px;margin-bottom:12px">
      <strong style="color:#065f46">Edit generated lesson plan</strong>
      <p style="font-size:.88rem;color:#047857;margin:4px 0 8px">Change any text below before download. Use <em>Edit mode</em> to type directly in the plan, or edit the structured fields.</p>
      <div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:8px">
        <button type="button" class="btn btn-gen" onclick="toggleEditMode()">✎ Edit mode on/off</button>
        <button type="button" class="btn btn-sample" onclick="openStructuredEdit()">Structured fields</button>
        <button type="button" class="btn btn-clear" onclick="reRenderFromLAST()">Reset view from data</button>
      </div>
      <div id="structuredEdit" style="display:none;margin-top:8px">
        <div class="form-grid">
          <div class="form-group full"><label>Key unit competence</label><textarea id="editKeyCompetence" rows="2"></textarea></div>
          <div class="form-group full"><label>Instructional objectives</label><textarea id="editObjective" rows="2"></textarea></div>
          <div class="form-group full"><label>Learning materials</label><textarea id="editMaterials" rows="2"></textarea></div>
          <div class="form-group full"><label>References</label><textarea id="editReferences" rows="2"></textarea></div>
          <div class="form-group full"><label>Introduction — Teacher (one line per bullet)</label><textarea id="editIntroT" rows="3"></textarea></div>
          <div class="form-group full"><label>Introduction — Learner</label><textarea id="editIntroL" rows="3"></textarea></div>
          <div class="form-group full"><label>Introduction — Competences</label><textarea id="editIntroC" rows="2"></textarea></div>
          <div class="form-group full"><label>Development — Teacher</label><textarea id="editDevT" rows="4"></textarea></div>
          <div class="form-group full"><label>Development — Learner</label><textarea id="editDevL" rows="4"></textarea></div>
          <div class="form-group full"><label>Development — Competences</label><textarea id="editDevC" rows="2"></textarea></div>
          <div class="form-group full"><label>Conclusion — Teacher</label><textarea id="editConcT" rows="2"></textarea></div>
          <div class="form-group full"><label>Conclusion — Learner</label><textarea id="editConcL" rows="2"></textarea></div>
          <div class="form-group full"><label>Conclusion — Competences</label><textarea id="editConcC" rows="2"></textarea></div>
          <div class="form-group full"><label>Teacher's self-evaluation</label><textarea id="editSelfEval" rows="2"></textarea></div>
        </div>
        <button type="button" class="btn btn-gen" style="margin-top:8px" onclick="applyStructuredEdit()">Apply edits to plan</button>
      </div>
    </div>
    <div class="dl-choice no-print">
      <strong>Download format:</strong>
      <label><input type="checkbox" id="dlPdf" checked> PDF</label>
      <label><input type="checkbox" id="dlWord" checked> Word</label>
      <label><input type="checkbox" id="dlExcel" checked> Excel</label>
      <label><input type="checkbox" id="dlCsv"> CSV</label>
      <div style="margin-top:8px">
        <button class="btn btn-gen" onclick="downloadSelected()">Download selected</button>
        <button class="btn btn-pdf" onclick="exportPDF()">PDF</button>
        <button class="btn btn-word" onclick="exportWord()">Word</button>
        <button class="btn btn-excel" onclick="exportExcel()">Excel</button>
        <button class="btn btn-csv" onclick="exportCSV()">CSV</button>
        <button class="btn btn-print" onclick="window.print()">Print</button>
      </div>
    </div>
    <div class="no-print" id="selfEvalBox" style="background:#fefce8;border:2px solid #ca8a04;border-radius:8px;padding:12px;margin-bottom:12px">
      <label style="font-weight:700;color:#854d0e;display:block;margin-bottom:6px">
        Teacher's self-evaluation (select an option — system completes words from unit &amp; lesson title)
      </label>
      <select id="selfEvalSelect" onchange="onSelfEvalSelect()" style="width:100%;padding:8px;border:2px solid #ca8a04;border-radius:6px;font-size:.95rem;margin-bottom:8px">
        <option value="">Select evaluation level…</option>
        <option value="excellent">Excellent — objectives fully achieved</option>
        <option value="good">Good — most objectives achieved</option>
        <option value="average">Average — partial achievement; need more practice</option>
        <option value="needs_improve">Needs improvement — reteach key points</option>
        <option value="mixed">Mixed — strong group work; weak individual practice</option>
        <option value="time">Time management issue — extend development next time</option>
      </select>
      <textarea id="selfEvalInput" rows="3" placeholder="Auto-filled from your selection (you may edit)"
        style="width:100%;padding:8px;border:2px solid #ca8a04;border-radius:6px;font-family:inherit;font-size:.95rem"></textarea>
      <div style="margin-top:8px">
        <button type="button" class="btn btn-gen" onclick="applySelfEvaluation()">Apply to lesson plan</button>
      </div>
    </div>
    <div id="lessonPlanContent" class="lesson-plan"></div>
  </div>
</div>

<script>
var LAST = null; // structured data for exports

const T = {
en: {
  lessonPlan:"LESSON PLAN", schoolName:"School Name", teacherName:"Teacher's Name",
  term:"Term", date:"Date", subject:"Subject", class:"Class", unitNo:"Unit N°", lessonNo:"Lesson N°",
  duration:"Duration", classSize:"Class size",
  specialNeeds:"Type of Special Educational Needs to be catered for in this lesson and number of learners in each category",
  unitTitle:"Unit title", keyCompetence:"Key Unit Competence", lessonTitle:"Title of the lesson",
  objective:"Instructional Objective", planClass:"Plan for this Class (location: in / outside)",
  materials:"Learning Materials (for all learners)", references:"References",
  timing:"Timing for each step", activities:"Description of teaching and learning activity",
  activityNote:"",
  teacherAct:"Teacher activities", learnerAct:"Learner activities",
  competences:"Generic competences and Cross cutting issues to be addressed + a short explanation",
  introduction:"Introduction", development:"Development of the lesson", conclusion:"Conclusion",
  conclusionNote:"",
  selfEval:"Teacher self-evaluation", none:"None", classroom:"Classroom",
  lessonOf:"of", minutes:"min"
},
rw: {
  lessonPlan:"IMBATA Y'ISOMO", schoolName:"Izina ry'ishuri", teacherName:"Amazina y'umwarimu",
  term:"Igihembwe", date:"Itariki", subject:"Inyigisho", class:"Umwaka wa", unitNo:"Umutwe wa", lessonNo:"Isomo rya",
  duration:"Igihe isomo rimara", classSize:"Umubare w'abanyeshuri",
  specialNeeds:"Abafite ibyo bagenerwa byihariye mu myigire no mu myigishirize n'umubare wabo",
  unitTitle:"Umutwe", keyCompetence:"Ubushobozi bw'ingenzi bugamijwe", lessonTitle:"Isomo",
  objective:"Intego ngenamukoro", planClass:"Imiterere y'aho isomo ribera",
  materials:"Imfashanyigisho", references:"Imyandiko n'ibitabo byifashishijwe",
  timing:"Ibice by'isomo + igihe", activities:"Gusobanura muri make igikorwa umwarimu n'umunyeshuri basabwa gukora",
  activityNote:"",
  teacherAct:"Ibikorwa by'umwarimu", learnerAct:"Ibikorwa by'umunyeshuri",
  competences:"Ubushobozi n'ingingo nsanganyamasomo (andika ubushobozi + igisobanuro kigufi kigaragaza uko buzagerwaho)",
  introduction:"Intangiriro", development:"Isomo nyirizina", conclusion:"Umusozo w'isomo",
  conclusionNote:"-Ikomatanya<br>-Isuzuma",
  selfEval:"Kwisuzuma (umurezi)", none:"Nta na kimwe", classroom:"Mu ishuri",
  lessonOf:"muri", minutes:"Iminota"
},
fr: {
  lessonPlan:"FICHE DE LEÇON", schoolName:"Nom de l'école", teacherName:"Nom de l'enseignant",
  term:"Trimestre", date:"Date", subject:"Branche", class:"Classe", unitNo:"Contexte N°", lessonNo:"Leçon N°",
  duration:"Durée", classSize:"Nombre d'apprenants",
  specialNeeds:"Types de besoins particuliers et nombre d'apprenants concernés",
  unitTitle:"Contexte", keyCompetence:"Compétence clé", lessonTitle:"Titre de la leçon",
  objective:"Objectif Opérationnel de la leçon", planClass:"Lieu de la leçon",
  materials:"Matériel didactique", references:"Références",
  timing:"Etapes de la leçon et durée", activities:"Description des activités d'enseignement et d'apprentissage",
  activityNote:"",
  teacherAct:"Activités de l'enseignant", learnerAct:"Activités des apprenants",
  competences:"Compétences génériques et thèmes transversaux + quelques explications",
  introduction:"Introduction", development:"Développement de la leçon", conclusion:"Conclusion",
  conclusionNote:"-Synthèse<br>-Évaluation",
  selfEval:"Auto évaluation", none:"Aucun", classroom:"Salle de classe",
  lessonOf:"sur", minutes:"minutes"
},
sw: {
  lessonPlan:"ANDALIO LA SOMO", schoolName:"Jina la shule", teacherName:"Jina la Mwalimu",
  term:"Muhula", date:"Tarehe", subject:"Somo", class:"Kidato", unitNo:"Mada", lessonNo:"Idadi ya masomo",
  duration:"Muda", classSize:"Idadi ya wanafunzi",
  specialNeeds:"Aina ya wanafunzi wenye mahitaji maalum ya kielimu watakaoshughulikiwa katika somo, idadi yao kulingana na mahitaji yao",
  unitTitle:"Mada", keyCompetence:"Uwezo uhitajiwao katika somo", lessonTitle:"Somo lenyewe",
  objective:"Malengo ya kujifunza", planClass:"Mahali ambapo somo litapofundishiwa",
  materials:"Vifaa au zana kwa wanafunzi wote", references:"Vitabu vya rejea",
  timing:"Muda na hatua", activities:"Mbinu za kufundishia na kujifunzia",
  activityNote:"",
  teacherAct:"Wajibu wa mwalimu", learnerAct:"Wajibu wa mwanafunzi",
  competences:"Tafakuri tunduizi na maswala mtambuka vielezewe + maelezo mafupi",
  introduction:"Utangulizi", development:"Somo lenyewe", conclusion:"Hitimisho",
  conclusionNote:"",
  selfEval:"Tathmini ya mwalimu", none:"Hakuna", classroom:"Darasani",
  lessonOf:"kati ya", minutes:"Dakika"
}
}

function fillSample() {
  document.getElementById('schoolName').value = 'BWENDA Primary School';
  document.getElementById('teacherName').value = 'SIBONIYO LAMBERT';
  document.getElementById('term').value = '2';
  document.getElementById('date').value = '2026-03-13';
  document.getElementById('level').value = 'Primary';
  onLevelChange();
  document.getElementById('class').value = 'P5';
  document.getElementById('classGroup').value = 'A';
  document.getElementById('subject').value = 'SET';
  onSubjectChange();
  document.getElementById('unitNo').value = '11';
  document.getElementById('lessonNo').value = '1';
  document.getElementById('totalLessons').value = '5';
  document.getElementById('duration').value = '40';
  document.getElementById('classSize').value = '43';
  document.getElementById('unitTitleSelect').value = 'Human digestive system';
  onUnitSelect();
  document.getElementById('unitTitle').value = 'Human digestive system';
  document.getElementById('lessonTitleSelect').value = 'Parts and functions of digestive system';
  document.getElementById('lessonTitle').value = 'Parts and functions of digestive system';
  document.getElementById('selfEvalSelect').value = 'good';
  onSelfEvalSelect();
  document.getElementById('keyCompetence').value = 'Identify the parts and functions of the digestive system, explain the stages of digestion, apply hygiene practices and show positive habits to keep their digestion healthy';
  document.getElementById('objective').value = '1. Using the guide provided in the textbook, P5 students will be able to identify different parts of the digestive system accurately.\n2. Using diagrams and models, P5 students will explain the functions of each digestive part clearly.';
  document.getElementById('specialNeeds').value = 'None';
  document.getElementById('location').value = 'Classroom';
  document.getElementById('materials').value = "SET P5 Learner's Book, 3D model of the digestive system, markers and chart paper.";
  document.getElementById('references').value = "Rwanda Education Board. (2025). SET Learner's Book. Kigali: REB.\nRwanda Education Board. (2025). Teacher's Guide for SET. Kigali: REB.";
  document.getElementById('errorMessage').textContent = '';
}

function clearForm() {
  document.getElementById('lessonForm').reset();
  document.getElementById('resultContainer').classList.remove('show');
  document.getElementById('errorMessage').textContent = '';
  LAST = null;
}

function esc(s) {
  return String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function bullets(arr) {
  return '<ul>' + (arr || []).map(function (x) { return '<li>- ' + esc(x) + '</li>'; }).join('') + '</ul>';
}

/** Generate rich, topic-specific activities like the sample PDF */

/* ===== SELECT CATALOGS (Level / Class / Subject / Units / Lessons) ===== */
var CLASS_BY_LEVEL = {
  'Primary': ['P1','P2','P3','P4','P5','P6'],
  'Ordinary Level': ['S1','S2','S3'],
  'Advanced Level': ['S4','S5','S6']
};
var SUBJECTS_BY_LEVEL = {
  'Primary': ['SET','Mathematics','English','Kinyarwanda','Social Studies','Music','Creative Arts','Physical Education','French','Religious Education'],
  'Ordinary Level': ['Mathematics','Physics','Chemistry','Biology','English','Kinyarwanda','French','History','Geography','ICT','Entrepreneurship','Kiswahili'],
  'Advanced Level': ['Mathematics','Physics','Chemistry','Biology','English','Economics','Geography','History','ICT','Entrepreneurship','General Paper']
};

/* REB CBC 2026-2027 content distribution / schemes of work — units & lessons by class + subject
   Source: elearning.reb.rw course 1806 + content distribution PDFs */
var UNITS_BY_SUBJECT_CLASS = {
  'SET': {
    'P1': [
      { unitNo:1, unit:'Natural and artificial things', lessons:['Natural things','Artificial (man-made) things','Differences between natural and artificial things','Caring for natural things','End of unit assessment'] },
      { unitNo:2, unit:'Making toys and teaching and learning materials', lessons:['Making a toy from paper','Making toys from local materials','Making simple teaching aids','Practical activities','End of unit assessment'] },
      { unitNo:3, unit:'Computer my friend', lessons:['Parts of a computer','Uses of computer parts','Taking care of a computer','Safe sitting posture','End of unit assessment'] },
      { unitNo:4, unit:'Plants', lessons:['Main parts of a plant','Importance of plant parts','Grouping plants by use','Caring for plants','End of unit assessment'] }
    ],
    'P2': [
      { unitNo:1, unit:'Natural and artificial things', lessons:['Natural things around us','Artificial things around us','Comparing natural and artificial things','Protecting natural things','End of unit assessment'] },
      { unitNo:2, unit:'Making toys and teaching and learning materials', lessons:['Making toys from paper','Making toys from local materials','Making teaching and learning materials','Practical activities','End of unit assessment'] },
      { unitNo:3, unit:'Computer my friend', lessons:['Main parts of a computer','Using a computer correctly and safely','Taking care of a computer','Practical activities','End of unit assessment'] },
      { unitNo:4, unit:'Plants', lessons:['Main parts of a plant','Importance of plant parts','Grouping plants by their use','Caring for plants','End of unit assessment'] }
    ],
    'P3': [
      { unitNo:1, unit:'Agricultural tools', lessons:['The common agricultural tools','Uses of agricultural tools','Maintenance of agricultural tools','Safe use of agricultural tools','End of unit assessment'] },
      { unitNo:2, unit:'Writing with a computer', lessons:['The keyboard and touchpad','Opening and saving a document','Writing capital and small letters','Practical activities on writing with a computer','End of unit assessment'] },
      { unitNo:3, unit:'Drinking water', lessons:['Preparation of drinking water by boiling','Preparation of drinking water by filtering','Storing drinking water safely','Hygiene of drinking water','End of unit assessment'] },
      { unitNo:4, unit:'Animals', lessons:['Vertebrate animals','Invertebrate animals','Differences between vertebrates and invertebrates','Caring for animals','End of unit assessment'] }
    ],
        'P4': [
      { unitNo:1, unit:'Objects production', keyCompetence:'Make useful and safe objects using local materials and take care of them properly.', lessons:['Making toys using wires','Making objects from local materials','Making learning materials','Safety when making objects','Caring for objects produced','End of unit assessment'] },
      { unitNo:2, unit:'ICT basics', keyCompetence:'Define different ICT terms used in the ICT domain and use them in simple tasks.', lessons:['Common ICT terms','Parts of a computer','Using a computer safely','Practical ICT activities','End of unit assessment'] },
      { unitNo:3, unit:'Graphics and multimedia', keyCompetence:'Use basic graphic and multimedia tools to create simple drawings and media products.', lessons:['Paint tools and freehand drawing','Drawing shapes','Working with multimedia files','Creativity with graphics','End of unit assessment'] },
      { unitNo:4, unit:'Programming for children', keyCompetence:'Design and construct geometric shapes in Turtle Art and create simple programs.', lessons:['Introduction to Turtle Art','Drawing geometric shapes','Simple programming commands','Practical projects','End of unit assessment'] },
      { unitNo:5, unit:'Air, wind and sound', keyCompetence:'Explain the importance of air, wind, and sound, their uses and how to protect them.', lessons:['Air and its importance','Wind and its effects','Sound and noise','Protecting air and reducing noise','End of unit assessment'] },
      { unitNo:6, unit:'Soil', keyCompetence:'Explain what soil is, describe its uses and importance, and show care for soil.', lessons:['What is soil','Composition of soil','Uses of soil','Soil erosion and prevention','Fertile soil','End of unit assessment'] },
      { unitNo:7, unit:'Animals classification', keyCompetence:'Classify animals according to how they breathe, move and live, and show care for animals.', lessons:['Animals that live on land','Animals that live in water','How animals breathe','How animals move','Caring for animals','End of unit assessment'] },
      { unitNo:8, unit:'Parts of a plant', keyCompetence:'Identify main parts of a plant and explain their functions.', lessons:['Root','Stem','Leaf','Flower','Fruit','End of unit assessment'] },
      { unitNo:10, unit:'Human sensory organs', keyCompetence:'Name the five sense organs, describe their functions and show care for them.', lessons:['The five sense organs','Functions of sense organs','Caring for sense organs','End of unit assessment'] },
      { unitNo:11, unit:'Human skeleton', keyCompetence:'Identify main bones of the human skeleton, explain their functions and show healthy habits for bones.', lessons:['Main bones of the skeleton','Functions of the skeleton','Caring for bones','End of unit assessment'] },
      { unitNo:12, unit:'Muscles', keyCompetence:'Name the main types of muscles, explain what muscles do and how to keep them healthy.', lessons:['Types of muscles','Functions of muscles','Keeping muscles healthy','End of unit assessment'] }
    ],

        'P5': [
      { unitNo:1, unit:'Carpentry tools', keyCompetence:'Name different carpentry tools, explain how to use and care for them, use them safely in simple tasks and show responsibility and interest in safe carpentry work.', lessons:['Identification of carpentry tools','Use and maintenance of carpentry tools','Dangers of carpentry tools and health and safety measures','End of unit assessment'] },
      { unitNo:2, unit:'Masonry tools', keyCompetence:'Name masonry tools, explain how to use and care for them, handle them safely during simple tasks and show responsibility in using and maintaining masonry tools.', lessons:['Identification of masonry tools','Use and care of masonry tools','Dangers of masonry tools and safety measures','End of unit assessment'] },
      { unitNo:3, unit:'Light', keyCompetence:'Describe how light moves and behaves, perform simple experiments to observe how light travels, reflects, and bends, and show interest by exploring light in their surroundings.', lessons:['Light propagation','Types of medium for light propagation (transparent, translucent and opaque)','Reflection of light','Refraction (bending) of light','Practical experiments on light','End of unit assessment'] },
      { unitNo:4, unit:'Computer memory and storage', keyCompetence:'Identify types of computer memory and storage devices, use them correctly and show responsibility in handling digital storage.', lessons:['Internal memory (ROM, RAM, hard disk)','External storage devices','Online storage','Memory size and units','End of unit assessment'] },
      { unitNo:5, unit:'Writing skills on computer', keyCompetence:'Insert table and pictures in a document.', lessons:['Typing and formatting text','Inserting a table in a document','Inserting pictures in a document','Editing rows and columns','End of unit assessment'] },
      { unitNo:6, unit:'Computer research', keyCompetence:'Use a computer and internet tools to identify reliable sources, search for information and apply the information appropriately in learning tasks.', lessons:['Terms used on the internet','Internet security','Searching for information','Navigating reliable sources','End of unit assessment'] },
      { unitNo:7, unit:'Programming for children', keyCompetence:'Use Turtle Art, Scratch, and robotics to solve math problems, draw shapes and patterns, tell digital stories and create fun projects using code.', lessons:['Perform arithmetic operations in Turtle Art','Drawing complex shapes in Turtle Art','Using paint editor in Scratch','Using sound and music in Scratch projects','Programming a robot','End of unit assessment'] },
      { unitNo:8, unit:'Water pollution and purification', keyCompetence:'Identify causes of water pollution, apply simple purification methods and show care for clean water in the environment.', lessons:['Causes of water pollution','Methods of water purification','Making a simple water filter','Protecting water sources','End of unit assessment'] },
      { unitNo:9, unit:'Soil and fertilizers', keyCompetence:'Differentiate organic and inorganic fertilizers, explain their use and show responsibility in caring for soil.', lessons:['Types of soil','Organic fertilizers','Inorganic fertilizers','Making and using compost','End of unit assessment'] },
      { unitNo:10, unit:'Plants and environment', keyCompetence:'Identify types of plants (food crops, cash crops, medicinal), explain their importance and show care for the environment through planting.', lessons:['Food crops and cash crops','Medicinal plants','Afforestation and protecting the environment','End of unit assessment'] },
      { unitNo:11, unit:'Human digestive system', keyCompetence:'Identify the parts and functions of the digestive system, explain the stages of digestion, apply hygiene practices and show positive habits to keep their digestive system healthy.', lessons:['Parts and functions of the digestive system','Stages of digestion','Hygiene of the digestive system','Balanced diet and nutritional deficiency diseases','End of unit assessment'] },
      { unitNo:12, unit:'Communicable diseases', keyCompetence:'Identify common communicable diseases, explain prevention measures and show responsibility for personal and community health.', lessons:['Waterborne diseases','Foodborne diseases','Prevention of communicable diseases','End of unit assessment'] },
      { unitNo:13, unit:'Electricity', keyCompetence:'Explain how electricity is made and why it is important, identify sources and parts of electric circuits, build and manage simple circuits, and show safe habits when working with electricity.', lessons:['Importance of electricity','Production of electricity','Sources of electricity','Simple electric circuits','Safety with electricity','End of unit assessment'] }
    ],

        'P6': [
      { unitNo:1, unit:'Mechanics tools', keyCompetence:'Use mechanics tools correctly, take care of them, and act safely and responsibly when working with tools.', lessons:['The common mechanics tools','Use and maintenance of mechanics tools','Dangers and safety precautions when using mechanics tools','Practical work on mechanics tools','End of unit assessment'] },
      { unitNo:2, unit:'Simple machines', keyCompetence:'Identify, use, and care for simple machines in a safe and responsible way, and understand how they help make work easier.', lessons:['Definition of simple machine and difference from other tools','Types of simple machines','Classes of levers','Practical use and care of simple machines','End of unit assessment'] },
      { unitNo:3, unit:'Air pollution', keyCompetence:'Identify causes and dangers of air pollution, take action to reduce it, and show care for clean and healthy air in the environment.', lessons:['Definition of air pollution and common air pollutants','Causes and dangers of air pollution','Ways to reduce air pollution','End of unit assessment and remediation'] },
      { unitNo:4, unit:'Cows and goats management', keyCompetence:'Explain how to care for cows and goats properly and show good habits in keeping animals that help families and communities.', lessons:['Characteristics of a good cowshed/goat shelter','Feeding cows and goats','Health and sanitation of cattle and goats','Parasitic and infectious diseases','Importance of cattle/goat farming','End of unit assessment'] },
      { unitNo:5, unit:'Word processing and spreadsheet', keyCompetence:'Manipulate documents using a Word Processing program, perform mathematical calculations using a Spreadsheet program and format a spreadsheet document.', lessons:['Making text bold, italic and underlined','Formatting paragraphs and pages','Creating and editing a spreadsheet','Performing calculations in a spreadsheet','Formatting a spreadsheet document','End of unit assessment'] },
      { unitNo:6, unit:'Emerging technologies', keyCompetence:'Identify emerging technologies (robotics, virtual reality, augmented reality) and use simple digital tools responsibly.', lessons:['Introduction to robotics','Virtual reality and augmented reality','Practical exploration of emerging technologies','End of unit assessment'] },
      { unitNo:7, unit:'Programming for children', keyCompetence:'Use Scratch to design shapes, do calculations, and create simple animations with sprites and backgrounds.', lessons:['Drawing polygons with Turtle Blocks','Creating sprites and backgrounds in Scratch','Simple animations and calculations','End of unit assessment'] },
      { unitNo:8, unit:'Plants reproduction', keyCompetence:'Identify parts of a flower, explain types of plant reproduction and show interest in growing plants at home and at school.', lessons:['Parts of a complete flower','Types of plant reproduction (sexual and asexual)','Growing plants from seeds and cuttings','End of unit assessment'] },
      { unitNo:9, unit:'Sustainable waste management', keyCompetence:'Understand different types of waste, use ways to manage waste safely, and show care for the environment by handling waste properly.', lessons:['Classification of wastes','Biodegradable and non-biodegradable wastes','Waste management methods (reduce, reuse, recycle, composting)','Safe disposal of hazardous waste','End of unit assessment'] },
      { unitNo:10, unit:'Human circulatory system', keyCompetence:'Identify parts of the circulatory system, describe how it works, and develop healthy habits to keep it working well.', lessons:['Main function of the circulatory system','Parts of the circulatory system','How blood circulates','Healthy habits for the circulatory system','End of unit assessment'] },
      { unitNo:11, unit:'Human respiratory system', keyCompetence:'Identify organs of the human respiratory system, explain how breathing works and show healthy respiratory habits.', lessons:['Organs of the human respiratory system and their main functions','The process of breathing','Diseases of the respiratory system and prevention','End of unit assessment'] },
      { unitNo:12, unit:'Human reproductive system', keyCompetence:'Identify the parts and functions of the human reproductive system, explain hygiene practices, puberty changes, and responsible behaviors, and show positive attitudes toward prevention of diseases and respect for people living with STIs and HIV.', lessons:['Functions of human reproductive system','Major internal and external parts of the male reproductive organ and their functions','Major female reproductive systems','Puberty changes and hygiene','Prevention of STIs and HIV','End of unit assessment'] },
      { unitNo:13, unit:'Energy management', keyCompetence:'Identify different forms and sources of energy, show how energy changes from one form to another, use energy wisely, and care about using clean energy to protect the environment.', lessons:['Definition of energy','Forms and sources of energy','Energy transformation','Using energy wisely and clean energy','End of unit assessment'] }
    ],

  },
  'Kinyarwanda': {
    'P6': [
      { unitNo:1, unit:'Umuco Nyarwanda', keyCompetence:'Gusoma no gusesengura imyandiko ijyanye n umuco nyarwanda; gukoresha neza ikeshamvugo n ibinyazina; guhanga indirimbo; gutandukanya amoko y insigamigani.', lessons:['Gusoma imyandiko','Gusesengura','Indirimbo','Isuzuma risoza umutwe'] },
      { unitNo:2, unit:'Ibidukikije', keyCompetence:'Gusoma no gusesengura imyandiko ku bidukikije.', lessons:['Gusoma','Gusesengura','Isuzuma risoza umutwe'] },
      { unitNo:3, unit:'Ubuzima', keyCompetence:'Gusoma no gusesengura imyandiko ku buzima.', lessons:['Gusoma','Gusesengura','Isuzuma risoza umutwe'] },
      { unitNo:4, unit:'Uburinganire n ubuzuzanye', keyCompetence:'Gusoma no gusesengura imyandiko ku buringanire n ubuzuzanye.', lessons:['Gusoma','Gusesengura','Isuzuma risoza umutwe'] },
      { unitNo:5, unit:'Itumanaho n ikoranabuhanga', keyCompetence:'Gusoma no gusesengura imyandiko ku itumanaho n ikoranabuhanga.', lessons:['Gusoma','Gusesengura','Isuzuma risoza umutwe'] },
      { unitNo:6, unit:'Ubukungu n ubukungu bw igihugu', keyCompetence:'Gusoma no gusesengura imyandiko ku bukungu.', lessons:['Gusoma','Gusesengura','Isuzuma risoza umutwe'] }
    ],

    'P5': [
      { unitNo:1, unit:'Umuco Nyarwanda', keyCompetence:'Gusoma no gusesengura imyandiko ijyanye n umuco Nyarwanda.', lessons:['Gusoma imyandiko','Gusesengura','Isuzuma risoza umutwe'] },
      { unitNo:2, unit:'Ibidukikije', keyCompetence:'Gusoma no gusesengura imyandiko ku bidukikije.', lessons:['Gusoma','Gusesengura','Isuzuma risoza umutwe'] },
      { unitNo:3, unit:'Uburenganzira bwa muntu', keyCompetence:'Gusoma no gusesengura imyandiko ku burenganzira bwa muntu.', lessons:['Gusoma','Gusesengura','Isuzuma risoza umutwe'] },
      { unitNo:4, unit:'Ubuzima', keyCompetence:'Gusoma no gusesengura imyandiko ku buzima.', lessons:['Gusoma','Gusesengura','Isuzuma risoza umutwe'] },
      { unitNo:5, unit:'Imiyoborere myiza', keyCompetence:'Gusoma no gusesengura imyandiko ku miyoborere myiza.', lessons:['Gusoma','Gusesengura','Isuzuma risoza umutwe'] },
      { unitNo:6, unit:'Itumanaho n ikoranabuhanga', keyCompetence:'Gusoma no gusesengura imyandiko ku itumanaho n ikoranabuhanga.', lessons:['Gusoma','Gusesengura','Isuzuma risoza umutwe'] }
    ],

        'P4': [
      { unitNo:1, unit:'Isubiramo', keyCompetence:'Gusubiramo ibyiganwe mu mwaka ushize no gukoresha ururimi mu mibiganiro ya buri munsi.', lessons:['Gusubiramo indimi n\u2019ibisobanuro','Gusoma n\u2019kwandika','Isuzuma risoza umutwe'] },
      { unitNo:2, unit:'Kubana neza n\u2019abandi', keyCompetence:'Gukoresha ikinyarwanda mu kubana neza n\u2019abandi no kugaragaza ubwisanzure.', lessons:['Imigenzo yo kubana neza','Gusoma inkuru','Kwandika','Isuzuma risoza umutwe'] },
      { unitNo:3, unit:'Umuco Nyarwanda', keyCompetence:'Gusoma no gusesengura inyandiko zijyanye n\u2019umuco Nyarwanda.', lessons:['Umuco Nyarwanda','Inkuru n\u2019imigenzo','Isuzuma risoza umutwe'] },
      { unitNo:4, unit:'Inkuru', keyCompetence:'Gusoma, gusesengura no guhanga inkuru ifite umutwe, intangiriro, igihimba n\u2019umusoso.', lessons:['Imiterere y\u2019inkuru','Gusoma inkuru','Guhanga inkuru','Isuzuma risoza umutwe'] },
      { unitNo:5, unit:'Imikino n\u2019imyidagaduro', keyCompetence:'Gukoresha ikinyarwanda mu mikino n\u2019imyidagaduro.', lessons:['Imikino','Imyidagaduro','Isuzuma risoza umutwe'] },
      { unitNo:6, unit:'Gukunda no kwitabira umurimo', keyCompetence:'Gukoresha ikinyarwanda mu kuvuga ku mirimo no kuyitabira.', lessons:['Imirimo','Kwitabira umurimo','Isuzuma risoza umutwe'] }
    ],
  

    'P1': [
      { unitNo:1, unit:'Guhimbarwa n\'ibisobanuro by\'ibanze', keyCompetence:'Gukoresha indimi z\'ubuvanganzo n\'ibisobanuro by\'ibanze mu miburo ya buri munsi.', lessons:['Guhimbarwa','Amazina y\'abantu n\'ibintu','Ibibazo by\'ibanze','Gusubiza neza','Isuzuma ry\'umwanya'] },
      { unitNo:2, unit:'Umuryango n\'urugo', keyCompetence:'Guvuga ku bagize umuryango n\'ibintu bihari mu rugo.', lessons:['Abagize umuryango','Imirimo y\'abagize umuryango','Ibintu bihari mu rugo','Gusabana mu rugo','Isuzuma ry\'umwanya'] },
      { unitNo:3, unit:'Ishuri n\'ibikoresho by\'ishuri', keyCompetence:'Guvuga ku ishuri n\'ibikoresho by\'ishuri ukoresheje interuro zoroheje.', lessons:['Amazina y\'ibikoresho by\'ishuri','Imirimo yo mu ishuri','Amabwiriza yo mu ishuri','Isuzuma ry\'umwanya'] },
      { unitNo:4, unit:'Amatungo n\'ibiribwa', keyCompetence:'Guvuga ku matungo n\'ibiribwa by\'ibanze.', lessons:['Amatungo yo mu rugo','Ibiribwa by\'ibanze','Ibyo nkunda / ntakunde','Isuzuma ry\'umwanya'] }
    ],
    'P2': [
      { unitNo:1, unit:'Guhimbarwa n\'imico myiza', keyCompetence:'Gukoresha indimi z\'ubuvanganzo n\'imico myiza mu miburo.', lessons:['Guhimbarwa mu buryo butandukanye','Imico myiza (murakoze, mbabarira)','Gusaba ubufasha','Isuzuma ry\'umwanya'] },
      { unitNo:2, unit:'Umuryango munini', keyCompetence:'Guvuga ku bagize umuryango munini n\'inshingano zabo.', lessons:['Abagize umuryango munini','Inshingano mu muryango','Kubana neza','Isuzuma ry\'umwanya'] },
      { unitNo:3, unit:'Ahantu n\'ibikorwa byo mu mudugudu', keyCompetence:'Guvuga ku hantu n\'ibikorwa byo mu mudugudu.', lessons:['Ahantu h\'ingenzi mu mudugudu','Abafasha mu mudugudu','Kubungabunga umudugudu','Isuzuma ry\'umwanya'] },
      { unitNo:4, unit:'Imyigishirize n\'imikino', keyCompetence:'Guvuga ku myigishirize n\'imikino y\'abana.', lessons:['Imyigishirize yo ku ishuri','Imikino y\'abana','Gukina neza','Isuzuma ry\'umwanya'] }
    ],
    'P3': [
      { unitNo:1, unit:'Kwiyumva n\'abandi', keyCompetence:'Kwiyumva no kuvuga ku bandi ukoresheje interuro zuzuye.', lessons:['Kwiyumva','Gusobanura abandi','Ingero z\'interuro','Isuzuma ry\'umwanya'] },
      { unitNo:2, unit:'Imirimo ya buri munsi', keyCompetence:'Guvuga ku mirimo ya buri munsi n\'igihe ikorwamo.', lessons:['Imirimo ya mu gitondo','Imirimo ya ku ishuri','Imirimo yo mu rugo','Isuzuma ry\'umwanya'] },
      { unitNo:3, unit:'Ubuzima n\'isuku', keyCompetence:'Guvuga ku buzima n\'isuku ukoresheje indimi zifatika.', lessons:['Isuku y\'umubiri','Ibiribwa byiza','Kwita ku buzima','Isuzuma ry\'umwanya'] },
      { unitNo:4, unit:'Imikino n\'imihango y\'umuco', keyCompetence:'Guvuga ku mikino n\'imihango y\'umuco nyarwanda.', lessons:['Imikino gakondo','Imihango y\'ingenzi','Kubaha umuco','Isuzuma ry\'umwanya'] }
    ],
},

  'French': {
    'P6': [
      { unitNo:1, unit:'Retour sur les contenus de 5e annee', keyCompetence:'Reviser et consolider les acquis de 5e annee.', lessons:['Vocabulaire temps et routine','Revision grammaire','Evaluation de l unite'] },
      { unitNo:2, unit:'Les droits de l enfant', keyCompetence:'Parler des droits de l enfant.', lessons:['Mise en route','Droits fondamentaux','Evaluation de l unite'] },
      { unitNo:3, unit:'Les pays et les nationalites', keyCompetence:'Parler des pays et des nationalites.', lessons:['Mise en route','Pays et nationalites','Evaluation de l unite'] },
      { unitNo:4, unit:'Les traditions nationales', keyCompetence:'Comprendre et evoquer les traditions nationales.', lessons:['Mise en route','Traditions','Evaluation de l unite'] },
      { unitNo:5, unit:'Fetes et anniversaires', keyCompetence:'Parler des fetes et des anniversaires.', lessons:['Mise en route','Fetes','Evaluation de l unite'] },
      { unitNo:6, unit:'Les loisirs et les sports', keyCompetence:'Parler des loisirs et des sports.', lessons:['Mise en route','Loisirs et sports','Evaluation de l unite'] },
      { unitNo:7, unit:'Les reseaux sociaux', keyCompetence:'Parler des reseaux sociaux.', lessons:['Mise en route','Reseaux sociaux','Evaluation de l unite'] },
      { unitNo:8, unit:'La nature et les animaux', keyCompetence:'Parler de la nature et des animaux.', lessons:['Mise en route','Animaux et nature','Evaluation de l unite'] },
      { unitNo:9, unit:'Mes projets futurs', keyCompetence:'Parler de mes projets futurs.', lessons:['Mise en route','Projets futurs','Evaluation de l unite'] }
    ],

    'P5': [
      { unitNo:1, unit:'Au marche', keyCompetence:'Faire des achats au marche.', lessons:['Mise en route','Vocabulaire du marche','Dialogues d achat','Evaluation de l unite'] },
      { unitNo:2, unit:'Se situer dans le temps', keyCompetence:'Situer des actions dans le temps.', lessons:['Mise en route','Heure et moments de la journee','Evaluation de l unite'] },
      { unitNo:3, unit:'La routine journaliere', keyCompetence:'Decrire sa routine journaliere.', lessons:['Mise en route','Verbes de routine','Evaluation de l unite'] },
      { unitNo:4, unit:'Mon emploi du temps', keyCompetence:'Parler de son emploi du temps.', lessons:['Mise en route','Activites de la semaine','Evaluation de l unite'] },
      { unitNo:5, unit:'Les dates et les mois de l annee', keyCompetence:'Parler des dates et des mois.', lessons:['Mise en route','Mois et dates','Evaluation de l unite'] },
      { unitNo:6, unit:'Mes activites futures', keyCompetence:'Parler de ses activites du futur.', lessons:['Mise en route','Futur proche','Evaluation de l unite'] },
      { unitNo:7, unit:'Gouts et preferences', keyCompetence:'Exprimer simplement ses gouts et ses preferences.', lessons:['Mise en route','J aime / Je n aime pas','Evaluation de l unite'] },
      { unitNo:8, unit:'Mes emotions et sentiments', keyCompetence:'Exprimer ses emotions et sentiments.', lessons:['Mise en route','Adjectifs d emotion','Evaluation de l unite'] },
      { unitNo:9, unit:'Raconter un evenement passe', keyCompetence:'Raconter des evenements passes.', lessons:['Mise en route','Passe compose simple','Evaluation de l unite'] }
    ],

    'P3': [
      { unitNo:1, unit:'Parler de soi et de son entourage', keyCompetence:'Communiquer avec des phrases simples pour se presenter, saluer, parler de sa famille et identifier les objets de la classe.', lessons:['Mise en route','Salutations','Se presenter','Famille','Objets de la classe','Evaluation de l unite'] },
      { unitNo:2, unit:'La langue de la classe', keyCompetence:'Utiliser les expressions de la langue de la classe.', lessons:['Mise en route','Consignes de classe','Evaluation de l unite'] },
      { unitNo:3, unit:'Se situer', keyCompetence:'Utiliser le lexique de localisation pour situer quelqu un ou quelque chose.', lessons:['Mise en route','Prepositions de lieu','Evaluation de l unite'] },
      { unitNo:4, unit:'La maison', keyCompetence:'Utiliser le vocabulaire de la maison.', lessons:['Mise en route','Pieces et objets','Evaluation de l unite'] },
      { unitNo:5, unit:'La meteo et le temps', keyCompetence:'Utiliser des expressions pour parler de la meteo et de l heure.', lessons:['Mise en route','Meteo','Heure','Evaluation de l unite'] },
      { unitNo:6, unit:'Les vetements', keyCompetence:'Decrire sa tenue vestimentaire.', lessons:['Mise en route','Noms des vetements','Couleurs','Evaluation de l unite'] }
    ],

    'P2': [
      { unitNo:1, unit:'Salutations', keyCompetence:'Utiliser des formules simples de salutation et de politesse pour entrer en contact.', lessons:['Mise en route','Bonjour Bonsoir Bonne nuit','Comment ca va','Evaluation de l unite'] },
      { unitNo:2, unit:'Se presenter', keyCompetence:'Se presenter et presenter une autre personne.', lessons:['Mise en route','Je m appelle','Il / Elle s appelle','Evaluation de l unite'] },
      { unitNo:3, unit:'Poser des questions', keyCompetence:'Poser des questions et repondre.', lessons:['Mise en route','Qui quoi quand','Evaluation de l unite'] },
      { unitNo:4, unit:'Presenter quelqu un', keyCompetence:'Presenter une personne (nom, metier, lieu).', lessons:['Mise en route','Il / Elle est','Profession','Evaluation de l unite'] },
      { unitNo:5, unit:'Ma famille', keyCompetence:'Nommer et presenter les membres de sa famille.', lessons:['Mise en route','Vocabulaire de la famille','Evaluation de l unite'] },
      { unitNo:6, unit:'Les objets de la classe', keyCompetence:'Identifier, nommer et utiliser les objets de la classe a l oral.', lessons:['Mise en route','Objets de la classe','Prepositions de lieu','Evaluation de l unite'] }
    ],

    'P1': [
      { unitNo:1, unit:'Salutations', keyCompetence:'Utiliser des formules simples de salutation et de politesse pour entrer en contact.', lessons:['Mise en route','Formules de salutation','Comment ca va','Prendre conge','Evaluation de l unite'] },
      { unitNo:2, unit:'Se presenter', keyCompetence:'Se presenter aux membres de la classe.', lessons:['Mise en route','Dire son nom','Presenter un camarade','Evaluation de l unite'] },
      { unitNo:3, unit:'Poser des questions', keyCompetence:'Poser des questions simples.', lessons:['Mise en route','Qui / quoi / quand','Evaluation de l unite'] },
      { unitNo:4, unit:'Presenter quelqu un', keyCompetence:'Etre capable de presenter quelqu un.', lessons:['Mise en route','Presenter une personne','Evaluation de l unite'] },
      { unitNo:5, unit:'Ma famille', keyCompetence:'Nommer et presenter les membres de sa famille.', lessons:['Mise en route','Membres de la famille','Evaluation de l unite'] },
      { unitNo:6, unit:'Les objets de la classe', keyCompetence:'Identifier, nommer et utiliser les objets de la classe.', lessons:['Mise en route','Objets de la classe','Localiser les objets','Evaluation de l unite'] }
    ],

        'P4': [
      { unitNo:1, unit:'Ce que nous avons appris en 3eme annee', keyCompetence:'Reconnaitre et reutiliser les mots et structures simples appris en 3eme annee (classe, maison, meteo, heure, vetements) pour communiquer a l oral et a l ecrit.', lessons:['Vocabulaire : langue de la classe et la maison','Vocabulaire : meteo et heure','Vocabulaire : vetements','Production orale et ecrite','Evaluation de l unite'] },
      { unitNo:2, unit:'De mon pays a mon quartier', keyCompetence:'Parler de son pays, de sa ville ou de son quartier en utilisant des prepositions de lieu.', lessons:['Vocabulaire du pays et du quartier','Prepositions de lieu','Production','Evaluation de l unite'] },
      { unitNo:3, unit:'Se situer et donner des indications', keyCompetence:'Comprendre et donner des indications simples pour se situer.', lessons:['Vocabulaire de localisation','Donner des indications','Production','Evaluation de l unite'] },
      { unitNo:4, unit:'Les moyens de transport', keyCompetence:'Utiliser le vocabulaire de base lie aux moyens de transport.', lessons:['Vocabulaire des moyens de transport','Parler de ses deplacements','Production','Evaluation de l unite'] },
      { unitNo:5, unit:'Au marche', keyCompetence:'Faire des achats au marche.', lessons:['Vocabulaire du marche','Dialogues d achat','Production','Evaluation de l unite'] },
      { unitNo:6, unit:'Les repas', keyCompetence:'Utiliser des expressions courantes liees aux repas.', lessons:['Vocabulaire des repas','Parler de ce que l on mange','Production','Evaluation de l unite'] },
      { unitNo:7, unit:'Le corps humain', keyCompetence:'Decrire le corps humain.', lessons:['Parties du corps','Decrire une personne','Production','Evaluation de l unite'] },
      { unitNo:8, unit:'La sante', keyCompetence:'Utiliser les expressions simples liees a la sante.', lessons:['Vocabulaire de la sante','Parler de la maladie et des soins','Production','Evaluation de l unite'] },
      { unitNo:9, unit:'Les professions', keyCompetence:'Utiliser le vocabulaire des professions et parler de son metier futur.', lessons:['Vocabulaire des metiers','Futur simple','Production','Evaluation de l unite'] }
    ],
  },

  'Music': {
    'P4': [
      { unitNo:1, unit:'Introduction to music', keyCompetence:'Pupils will be able to sing respecting the elements of music.', lessons:['The meaning and role of music','Elements of music (lyrics, melody, rhythm)','Singing songs about peace and friendship','End of unit assessment'] },
      { unitNo:2, unit:'Rhythm and beat', keyCompetence:'Pupils will be able to clap and move to the beat of simple songs.', lessons:['Feeling the beat','Clapping rhythms','Moving to music','End of unit assessment'] },
      { unitNo:3, unit:'Introduction to C scale', keyCompetence:'Pupils will be able to sing and recognize notes of the C scale.', lessons:['Notes of the C scale','Singing the C scale','Practice songs in C scale','End of unit assessment'] },
      { unitNo:4, unit:'Practice songs in C scale', keyCompetence:'Pupils will be able to perform simple songs in C scale with correct pitch and rhythm.', lessons:['Learning a song in C scale','Group performance','End of unit assessment'] }
    ],
  

    'P1': [
      { unitNo:1, unit:'Sounds around us', keyCompetence:'Identify and imitate environmental and body sounds.', lessons:['Sounds in the environment','Body percussion','Loud and soft sounds','End of unit assessment'] },
      { unitNo:2, unit:'Simple songs and rhymes', keyCompetence:'Sing simple songs and recite rhymes with actions.', lessons:['Action songs','Nursery rhymes','Singing in a group','End of unit assessment'] }
    ],
    'P2': [
      { unitNo:1, unit:'Rhythm and beat', keyCompetence:'Clap and move to a steady beat and simple rhythms.', lessons:['Steady beat','Simple rhythmic patterns','Clapping games','End of unit assessment'] },
      { unitNo:2, unit:'Singing and movement', keyCompetence:'Sing simple songs with coordinated movements.', lessons:['Call and response songs','Songs with movement','Group singing','End of unit assessment'] }
    ],
    'P3': [
      { unitNo:1, unit:'Introduction to pitch and melody', keyCompetence:'Differentiate high and low sounds and sing simple melodies.', lessons:['High and low sounds','Simple melodic patterns','Singing scales (intro)','End of unit assessment'] },
      { unitNo:2, unit:'Traditional songs', keyCompetence:'Sing and appreciate simple Rwandan traditional songs.', lessons:['Selected traditional songs','Meaning of the songs','Group performance','End of unit assessment'] }
    ],
    'P5': [
      { unitNo:1, unit:'Rhythm patterns and time signatures', keyCompetence:'Read and perform simple rhythm patterns in basic time signatures.', lessons:['2/4 and 3/4 time','Rhythm notation intro','Performing patterns','End of unit assessment'] },
      { unitNo:2, unit:'Melody and scales', keyCompetence:'Sing and recognise melodies based on major scales.', lessons:['Major scale review','Simple sight singing','Creating short melodies','End of unit assessment'] },
      { unitNo:3, unit:'Traditional and modern songs', keyCompetence:'Perform traditional and modern songs with expression.', lessons:['Rwandan traditional songs','Modern children songs','Expression and dynamics','End of unit assessment'] }
    ],
    'P6': [
      { unitNo:1, unit:'Music notation and reading', keyCompetence:'Read simple music notation and perform from score.', lessons:['Notes and rests','Simple melodies from notation','Rhythm reading','End of unit assessment'] },
      { unitNo:2, unit:'Harmony and part singing', keyCompetence:'Sing simple two-part songs and appreciate harmony.', lessons:['Singing in unison and parts','Simple rounds','Listening to harmony','End of unit assessment'] },
      { unitNo:3, unit:'Composition and performance', keyCompetence:'Create and perform short musical pieces in a group.', lessons:['Creating a short song or rhythm piece','Group rehearsal','Class performance','Reflection and improvement','End of unit assessment'] }
    ],
},

  'Physical Education': {
    'P3': [
      { unitNo:1, unit:'Motor control', keyCompetence:'Practice simple gymnastics moves carefully and smoothly.', lessons:['Introduction to rolling forward and backward exercises','Flexibility exercises and games','End of unit practical assessment'] },
      { unitNo:2, unit:'Movement and coordination', keyCompetence:'Move whole body in different directions and speeds with balance and control.', lessons:['Moving in different directions','Games for coordination','End of unit practical assessment'] },
      { unitNo:3, unit:'Athletic exercises and games', keyCompetence:'Practice short-distance running, jumping and throwing through fun exercises and games.', lessons:['Running exercises','Jumping exercises','Throwing exercises','End of unit practical assessment'] },
      { unitNo:4, unit:'Ball games played with our feet', keyCompetence:'Play football by practicing basic skills like passing and shooting the ball.', lessons:['Passing the ball','Shooting','Simple football games','End of unit practical assessment'] },
      { unitNo:5, unit:'Ball games primarily played with our hands', keyCompetence:'Move, pass and catch the ball while playing games like Volleyball, Basketball, Handball, Netball.', lessons:['Passing and catching','Moving with the ball','Simple hand-ball games','End of unit practical assessment'] },
      { unitNo:6, unit:'Health and well-being', keyCompetence:'Learn how to stay healthy by keeping clean, eating well, staying safe and active.', lessons:['Personal hygiene','Healthy habits','Safety in play','End of unit practical assessment'] }
    ],

    'P4': [
      { unitNo:1, unit:'Motor control', keyCompetence:'Perform a range of aerobic warm-up, stretching and relaxation exercises and develop basic Gymnastics skills using good body control.', lessons:['Importance of physical activities and sports to health','Aerobic warm-up exercises','Stretching exercises','Relaxation exercises','Basic gymnastics skills','End of unit practical assessment'] },
      { unitNo:2, unit:'Body control and balance', keyCompetence:'Practise simple Gymnastics moves correctly, using the body well.', lessons:['Body control exercises','Balance exercises','Simple gymnastics moves','End of unit assessment and remediation'] },
      { unitNo:3, unit:'Rwandan traditional games', keyCompetence:'Play, enjoy and appreciate Rwandan traditional games.', lessons:['Introduction to traditional games','Playing selected traditional games','End of unit practical assessment'] },
      { unitNo:4, unit:'Introduction to Gymnastics', keyCompetence:'Perform floor exercises and use simple Gymnastic equipment safely.', lessons:['Floor exercises','Using simple gymnastic equipment','End of unit practical assessment'] },
      { unitNo:5, unit:'Introduction to Athletics', keyCompetence:'Practise basic throwing, running and jumping skills correctly.', lessons:['Basic running skills','Basic jumping skills','Basic throwing skills','End of unit practical assessment'] },
      { unitNo:6, unit:'Introduction to Football', keyCompetence:'Practise basic football skills (kick, pass, control) correctly.', lessons:['Kicking and shooting','Passing the ball','Controlling the ball','End of unit practical assessment'] },
      { unitNo:7, unit:'Introduction to Volleyball', keyCompetence:'Practice and use simple volleyball skills like serving, passing and receiving.', lessons:['Serving','Passing and receiving','Simple game situations','End of unit practical assessment'] },
      { unitNo:8, unit:'Introduction to Basketball', keyCompetence:'Practise dribbling, passing and shooting the ball in Basketball games.', lessons:['Dribbling','Passing','Shooting','End of unit practical assessment'] },
      { unitNo:9, unit:'Introduction to Handball', keyCompetence:'Practice basic Handball skills like passing, catching and shooting during the game.', lessons:['Passing and catching','Shooting','Simple game play','End of unit practical assessment'] }
    ],
  

    'P1': [
      { unitNo:1, unit:'Body awareness and basic movements', keyCompetence:'Perform basic body movements with control and awareness.', lessons:['Body parts and movement','Walking, running, jumping, hopping','Stopping and starting on signal','Simple stretching','End of unit assessment'] },
      { unitNo:2, unit:'Locomotor and non-locomotor skills', keyCompetence:'Demonstrate locomotor and non-locomotor skills in games and activities.', lessons:['Walking and running games','Jumping and hopping games','Bending, stretching, twisting','Balancing on one foot','End of unit assessment'] },
      { unitNo:3, unit:'Simple games and play', keyCompetence:'Participate in simple cooperative games and follow basic rules.', lessons:['Circle games','Relay races (simple)','Throwing and catching soft objects','Playing fairly','End of unit assessment'] },
      { unitNo:4, unit:'Health and safety in physical activity', keyCompetence:'Practise safe behaviour and basic hygiene related to physical activity.', lessons:['Warm-up and cool-down','Safe use of space','Personal hygiene after activity','Drinking water','End of unit assessment'] }
    ],
    'P2': [
      { unitNo:1, unit:'Basic motor skills', keyCompetence:'Perform fundamental motor skills with increasing control and coordination.', lessons:['Running and changing direction','Jumping for height and distance','Hopping and skipping','Galloping and sliding','End of unit assessment'] },
      { unitNo:2, unit:'Ball skills introduction', keyCompetence:'Handle soft balls with hands and feet in simple activities.', lessons:['Rolling and stopping a ball','Throwing underarm','Catching with two hands','Kicking a stationary ball','End of unit assessment'] },
      { unitNo:3, unit:'Simple traditional and cooperative games', keyCompetence:'Play simple Rwandan traditional and cooperative games following rules.', lessons:['Selected traditional games','Cooperative games','Following rules','Team spirit','End of unit assessment'] },
      { unitNo:4, unit:'Fitness and healthy habits', keyCompetence:'Participate in activities that develop fitness and healthy habits.', lessons:['Simple strength activities','Flexibility exercises','Importance of regular activity','Rest and sleep','End of unit assessment'] }
    ],
    'P5': [
      { unitNo:1, unit:'Gymnastics and body control', keyCompetence:'Perform basic gymnastics movements with control, balance and safety.', lessons:['Forward and backward rolls','Balances and supports','Jumps and landings','Simple sequences','Safety in gymnastics','End of unit assessment'] },
      { unitNo:2, unit:'Athletics skills', keyCompetence:'Perform basic running, jumping and throwing techniques.', lessons:['Sprint starts and running form','Relay baton exchange','Long jump approach','Basic throws','End of unit assessment'] },
      { unitNo:3, unit:'Football and netball basics', keyCompetence:'Apply basic skills and rules of football and netball in small-sided games.', lessons:['Football: passing and receiving','Football: dribbling and shooting','Netball: passing and footwork','Small-sided games and rules','End of unit assessment'] },
      { unitNo:4, unit:'Traditional games and dance', keyCompetence:'Perform Rwandan traditional games and simple dance movements.', lessons:['Selected traditional games','Rhythmic movements','Simple traditional dance steps','Group performance','End of unit assessment'] },
      { unitNo:5, unit:'Health, fitness and fair play', keyCompetence:'Explain benefits of physical activity and demonstrate fair play and respect.', lessons:['Components of fitness','Warm-up and cool-down importance','Fair play and respect','Personal fitness goals','End of unit assessment'] }
    ],
    'P6': [
      { unitNo:1, unit:'Advanced gymnastics and sequences', keyCompetence:'Perform linked gymnastics movements and short sequences with good form.', lessons:['Rolls, balances and jumps in combination','Creating a short sequence','Peer observation and feedback','Safety and spotting','End of unit assessment'] },
      { unitNo:2, unit:'Athletics events', keyCompetence:'Demonstrate improved technique in selected running, jumping and throwing events.', lessons:['Sprints and relays','Long jump and high jump (basic)','Shot put or soft javelin','Event rules and measurement','End of unit assessment'] },
      { unitNo:3, unit:'Team sports – football, netball, volleyball', keyCompetence:'Apply skills, tactics and rules of selected team sports in game situations.', lessons:['Football skills and small-sided games','Netball skills and rules','Volleyball: dig, set, serve','Team tactics and fair play','End of unit assessment'] },
      { unitNo:4, unit:'Traditional sports and cultural dance', keyCompetence:'Perform and appreciate Rwandan traditional sports and cultural dance.', lessons:['Traditional sports and games','Cultural dance sequences','Group choreography','Performance and respect for culture','End of unit assessment'] },
      { unitNo:5, unit:'Fitness, health and leadership', keyCompetence:'Plan simple fitness activities and show leadership and responsibility in PE.', lessons:['Fitness testing (simple)','Designing a warm-up','Leading a small group','Lifelong physical activity','End of unit assessment'] }
    ],
},

  'English': {
    'P6': [
      { unitNo:1, unit:'Leisure and sports', keyCompetence:'Communicate effectively about leisure activities, music, and sports while promoting teamwork, respect for others preferences, and healthy use of leisure time.', lessons:['Talking about what you like doing','Music and leisure','Sports activities','Writing about leisure','End of unit assessment'] },
      { unitNo:2, unit:'Our environment', keyCompetence:'Communicate effectively about the environment and how to care for it.', lessons:['Talking about the environment','Protecting nature','End of unit assessment'] },
      { unitNo:3, unit:'Weather', keyCompetence:'Communicate effectively about weather and seasons.', lessons:['Describing the weather','Seasons','End of unit assessment'] },
      { unitNo:4, unit:'Behaviour, rules and laws', keyCompetence:'Communicate effectively about behaviour, rules, laws, and consequences while promoting respect and responsibility.', lessons:['Rules at school and home','Laws in the community','Good behaviour','End of unit assessment'] },
      { unitNo:5, unit:'Family relationships', keyCompetence:'Communicate effectively about family relationships and roles.', lessons:['Family members and roles','Relationships','End of unit assessment'] }
    ],

    'P4': [
      { unitNo:1, unit:'Our school', keyCompetence:'Communicate effectively about school life, classroom routines and people at school.', lessons:['Talking about our school','Classroom objects and routines','People at school','Writing about school','End of unit assessment'] },
      { unitNo:2, unit:'My family and home', keyCompetence:'Communicate effectively about family members and home activities.', lessons:['Family members','Home activities','Writing about family','End of unit assessment'] },
      { unitNo:3, unit:'Food and meals', keyCompetence:'Communicate effectively about food, meals and healthy eating.', lessons:['Naming foods','Talking about meals','Healthy eating habits','End of unit assessment'] },
      { unitNo:4, unit:'Places in our community', keyCompetence:'Communicate effectively about places in the community and directions.', lessons:['Places in the community','Giving directions','Writing about a place','End of unit assessment'] },
      { unitNo:5, unit:'Jobs and roles in home and community', keyCompetence:'Communicate effectively about jobs and roles at home and in the community while promoting responsibility and respect.', lessons:['Jobs at home','Jobs in the community','Roles and responsibilities','End of unit assessment'] }
    ],

    'P5': [
      { unitNo:1, unit:'Past and future events', keyCompetence:'Communicate effectively about past experiences and plans while promoting truthfulness in recounting events and responsibility in planning future actions.', lessons:['Reading comprehension about Amani\'s life story','Vocabulary for life stories','Telling your life story (speaking)','Writing about past experiences','Talking about future plans using going to / want to / would like to','Sounds and spelling','End of unit assessment'] },
      { unitNo:2, unit:'Classroom language', keyCompetence:'Communicate effectively during classroom interactions using appropriate language, respectful turn-taking, and cooperation in group work.', lessons:['Giving orders (reading comprehension)','Speaking — classroom instructions','Vocabulary for classroom language','Expressing different opinions and preferences','Writing classroom dialogues','Sounds and spelling','End of unit assessment'] },
      { unitNo:3, unit:'Reading for enjoyment', keyCompetence:'Communicate effectively about stories and reading habits while promoting a love of reading and respect for others\' preferences.', lessons:['Reading the story Eagle and Fox','Speaking about the story','Saying what I like reading','Talking about reading habits','Skimming and scanning','Writing about reading preferences','End of unit assessment'] },
      { unitNo:4, unit:'Describing places and wildlife in Rwanda', keyCompetence:'Communicate effectively about places, physical features and wildlife in Rwanda while promoting environmental care and respect.', lessons:['Describing Rwanda\'s map and landmarks','Describing uses of physical features','Describing wildlife in Rwanda','Writing using passive voice','Sounds and spelling','End of unit assessment'] },
      { unitNo:5, unit:'Measurements', keyCompetence:'Communicate effectively about measurements, dimensions, height and size while promoting accuracy and cooperation.', lessons:['Saying and asking about dimensions','Comparing sizes','Discussing people\'s height and size','Writing measurement descriptions','End of unit assessment'] },
      { unitNo:6, unit:'Transport', keyCompetence:'Communicate effectively about means of transport and journeys while promoting safety and appreciation of transport services.', lessons:['Talking about means of transport','Travel by + vocabulary','Describing places visited in Rwanda','Writing about journeys','Sounds and spelling','End of unit assessment'] },
      { unitNo:7, unit:'Hygiene and health', keyCompetence:'Communicate effectively about diet, hygiene and health while promoting healthy living, self-care, and community responsibility.', lessons:['Describing diet','Describing food groups and a balanced diet','Healthy hygiene practices','Talking about illness','Writing health advice using should','End of unit assessment'] },
      { unitNo:8, unit:'Crafts in Rwanda', keyCompetence:'Communicate effectively about traditional and modern crafts while promoting cultural heritage and creativity.', lessons:['Describing what people made in the past','Describing what people used tools for','Describing modern handcrafts','Describing how to make craft objects','End of unit assessment'] },
      { unitNo:9, unit:'Agriculture in Rwanda', keyCompetence:'Communicate effectively about farming tools, farming in the past and future improvements in agriculture.', lessons:['Describing farming tools in the past','Describing farming in the past','Talking about future improvements in agriculture','Writing about agriculture using future simple','End of unit assessment'] },
      { unitNo:10, unit:'Geography of the world', keyCompetence:'Communicate effectively about continents, countries and capital cities while promoting global awareness and respect for diversity.', lessons:['Naming continents, countries, and capital cities','Describing the position of countries','Talking about possibilities using can','Writing about countries','Sounds and spelling','End of unit assessment'] }
    ],
    'P1': [
      { unitNo:1, unit:'Greetings and introductions', keyCompetence:'Use simple greetings and introduce oneself and others in English.', lessons:['Hello, good morning, good afternoon','My name is...','This is my friend','Saying goodbye','End of unit assessment'] },
      { unitNo:2, unit:'My body and clothes', keyCompetence:'Name parts of the body and common clothes and use simple descriptive words.', lessons:['Parts of the body','Action words (clap, jump, sit)','Clothes we wear','Colours of clothes','End of unit assessment'] },
      { unitNo:3, unit:'My family and home', keyCompetence:'Talk about family members and rooms/objects at home using simple sentences.', lessons:['Family members','Rooms in the house','Objects at home','This is / These are','End of unit assessment'] },
      { unitNo:4, unit:'School and classroom', keyCompetence:'Name classroom objects and follow simple classroom instructions in English.', lessons:['Classroom objects','School places','Classroom instructions','Polite requests','End of unit assessment'] },
      { unitNo:5, unit:'Animals and food', keyCompetence:'Name common animals and foods and express likes in simple sentences.', lessons:['Domestic animals','Wild animals','Fruits and foods','I like / I do not like','End of unit assessment'] }
    ],
    'P2': [
      { unitNo:1, unit:'Greetings, introductions and courtesy', keyCompetence:'Use a wider range of greetings and polite expressions in daily situations.', lessons:['Formal and informal greetings','Introducing family and friends','Please, thank you, sorry, excuse me','Asking and answering How are you?','End of unit assessment'] },
      { unitNo:2, unit:'My day and daily routines', keyCompetence:'Describe simple daily routines using present simple.', lessons:['Times of the day','Daily activities','I wake up / I go to school','Sequencing with first, then, next','End of unit assessment'] },
      { unitNo:3, unit:'Food, meals and shopping', keyCompetence:'Talk about meals, food preferences and simple shopping dialogues.', lessons:['Meals of the day','Food items','I like / I prefer','At the market (simple dialogue)','End of unit assessment'] },
      { unitNo:4, unit:'Places in the community', keyCompetence:'Name and describe places in the community and give simple directions.', lessons:['Places in the village/town','Where is the... ?','Left, right, straight','Asking for and giving directions','End of unit assessment'] },
      { unitNo:5, unit:'Animals, plants and the environment', keyCompetence:'Describe animals and plants and talk about caring for the environment.', lessons:['Domestic and wild animals','Parts of a plant','Caring for animals and plants','Keeping our environment clean','End of unit assessment'] }
    ],
    'P3': [
      { unitNo:1, unit:'Talking about myself and others', keyCompetence:'Describe oneself and others using simple present and adjectives.', lessons:['Personal information','Physical appearance','Character adjectives','He/She is... / They are...','End of unit assessment'] },
      { unitNo:2, unit:'Daily routines and free time', keyCompetence:'Talk about daily routines and free-time activities using frequency words.', lessons:['Daily routines','Free-time activities','Always, usually, sometimes, never','Writing a short paragraph about my day','End of unit assessment'] },
      { unitNo:3, unit:'Food, health and hygiene', keyCompetence:'Discuss healthy food, hygiene practices and simple health advice.', lessons:['Healthy and unhealthy food','Personal hygiene','At the clinic (simple dialogue)','Giving advice: You should...','End of unit assessment'] },
      { unitNo:4, unit:'Our community and helpers', keyCompetence:'Describe community helpers and their roles using present simple.', lessons:['Community helpers','What does a... do?','Places of work','Thanking community helpers','End of unit assessment'] },
      { unitNo:5, unit:'Stories, songs and poems', keyCompetence:'Listen to and retell short stories and recite simple poems and songs.', lessons:['Listening to a short story','Retelling with pictures','Simple poems','Action songs','End of unit assessment'] }
    ],
},
  'Mathematics': {
    'P1': [
      { unitNo:1, unit:'Numbers from 1 to 5', lessons:['Introduction of the unit','Counting numbers 1 to 5','Reading and writing numbers 1 to 5','Ordering and comparing numbers 1 to 5','Addition of numbers within 5','Subtraction of numbers within 5','End of unit assessment'] },
      { unitNo:2, unit:'Numbers from 1 to 9', lessons:['Introduction of the unit','Counting numbers 1 to 9','Reading and writing numbers 1 to 9','Ordering and comparing','Addition and subtraction within 9','End of unit assessment'] },
      { unitNo:3, unit:'Numbers from 0 to 20', lessons:['Counting 0 to 20','Place value of tens and ones','Ordering and comparing','Addition and subtraction within 20','End of unit assessment'] },
      { unitNo:4, unit:'Numbers from 0 to 50', lessons:['Counting 0 to 50','Reading and writing numbers','Comparing and ordering','Simple word problems','End of unit assessment'] },
      { unitNo:8, unit:'Measuring length less than or equal to 10 m', lessons:['Comparing lengths','Non-standard units of length','The metre (m)','Measuring and comparing lengths','Addition and subtraction of lengths','End of unit assessment'] },
      { unitNo:9, unit:'Rwandan currency from 1 Frw to 100 Frw', lessons:['Recognising coins and notes','Counting money','Simple addition of money','Simple subtraction of money','Word problems on money','End of unit assessment'] },
      { unitNo:10, unit:'Location of objects and lines', lessons:['Location of objects on a line','Types of lines','Practical activities','End of unit assessment'] },
      { unitNo:11, unit:'Right angle, square and rectangle', lessons:['Identifying a right angle','Square','Rectangle','Drawing squares and rectangles','End of unit assessment'] }
    ],
    'P2': [
      { unitNo:1, unit:'Whole numbers from 0 up to 200', lessons:['Count, read and write numbers 0–200','Place value of each digit','Comparison of numbers 0–200','Ordering numbers','End of unit assessment'] },
      { unitNo:2, unit:'Whole numbers from 0 up to 500', lessons:['Count, read and write numbers 0–500','Place value','Comparison and ordering','Addition and subtraction','End of unit assessment'] }
    ],
        'P4': [
      { unitNo:1, unit:'Mathematical operations on whole numbers up to 100,000', keyCompetence:'Pupils will be able to read, write, compare, and make calculations on whole numbers up to 100,000.', lessons:['Introduction to the content of the unit','Place values of whole numbers up to 100,000','Reading and writing numbers up to 100,000','Comparing and ordering numbers','Addition of whole numbers','Subtraction of whole numbers','Multiplication of whole numbers','Division of whole numbers','Word problems','End of unit assessment'] },
      { unitNo:2, unit:'Positive and negative integers', keyCompetence:'Pupils will be able to solve problems involving positive and negative integers.', lessons:['Introduction to integers','Positive and negative numbers','Comparing integers','Operations on integers','End of unit assessment'] },
      { unitNo:3, unit:'Classifying numbers by their properties', keyCompetence:'Pupils will be able to classify numbers by their properties (even, odd, prime, composite, factors, multiples).', lessons:['Introduction to classification of numbers','Even and odd numbers','Prime and composite numbers','Factors and multiples','End of unit assessment'] },
      { unitNo:4, unit:'Fractions with the same denominator', keyCompetence:'Pupils will be able to add, subtract and compare fractions with the same denominator.', lessons:['Introduction to fractions','Proper fractions','Comparing like fractions','Addition of like fractions','Subtraction of like fractions','Word problems on fractions','End of unit assessment'] },
      { unitNo:5, unit:'Decimal numbers', keyCompetence:'Pupils will be able to add, subtract and compare decimal numbers.', lessons:['Introduction to decimals','Reading and writing decimals','Comparing decimals','Addition of decimals','Subtraction of decimals','End of unit assessment'] },
      { unitNo:6, unit:'Length measurements', keyCompetence:'Pupils will be able to convert between units of length and solve related problems.', lessons:['Introduction to length','Non-standard units of length','Standard units (mm, cm, m, km)','Conversion of length units','Perimeter','End of unit assessment'] },
      { unitNo:7, unit:'Capacity measurements', keyCompetence:'Pupils will be able to convert between units of capacity and solve related problems.', lessons:['Introduction to capacity','Units of capacity (ml, cl, l)','Conversion of capacity units','Addition and subtraction of capacity','End of unit assessment'] },
      { unitNo:8, unit:'Mass measurements', keyCompetence:'Pupils will be able to convert between units of mass and solve related problems.', lessons:['Introduction to mass','Units of mass (g, kg, tonne)','Conversion of mass units','Problems involving mass','End of unit assessment'] },
      { unitNo:9, unit:'Area and land measurements', keyCompetence:'Pupils will be able to use square units and calculate area of land and shapes.', lessons:['Concept of unit area','Units of area (cm², m², ha)','Conversion of area units','Calculating area','End of unit assessment'] },
      { unitNo:10, unit:'Time', keyCompetence:'Pupils will be able to tell, write, and convert time appropriately.', lessons:['Introduction to time','Reading clocks','AM and PM','Converting units of time','End of unit assessment'] },
      { unitNo:11, unit:'Number patterns', keyCompetence:'Pupils will be able to describe and generate number patterns.', lessons:['Introduction to patterns','Describing number sequences','Generating patterns','End of unit assessment'] },
      { unitNo:12, unit:'Solving problems involving missing numbers', keyCompetence:'Pupils will be able to solve problems involving missing numbers in addition and subtraction.', lessons:['Introduction to missing numbers','Missing numbers in addition','Missing numbers in subtraction','End of unit assessment'] },
      { unitNo:13, unit:'Lines and angles', keyCompetence:'Pupils will be able to use geometric properties of lines and angles.', lessons:['Types of lines','Types of angles','Drawing lines and angles','End of unit assessment'] },
      { unitNo:14, unit:'Shapes in the plane', keyCompetence:'Pupils will be able to classify triangles and quadrilaterals based on their properties.', lessons:['Naming triangles','Naming quadrilaterals','Properties of plane shapes','End of unit assessment'] },
      { unitNo:15, unit:'Area of shapes in the plane', keyCompetence:'Pupils will be able to use rectangles to determine area of shapes in the plane.', lessons:['Area of rectangles','Area of other plane shapes','End of unit assessment'] },
      { unitNo:16, unit:'Introduction to data handling', keyCompetence:'Pupils will be able to collect, represent and interpret simple data.', lessons:['Collecting data','Tables and tallies','Simple graphs','Interpreting data','End of unit assessment'] },
      { unitNo:17, unit:'Introduction to games of chance', keyCompetence:'Pupils will be able to play games of chance and discuss fairness.', lessons:['Chance and luck','Fair and unfair games','End of unit assessment'] }
    ],

  

    'P3': [
      { unitNo:1, unit:'Whole numbers from 0 up to 1000', keyCompetence:'Read, write, order and perform operations on whole numbers up to 1000.', lessons:['Counting and writing numbers up to 1000','Place value (hundreds, tens, ones)','Comparing and ordering numbers','Addition of numbers up to 1000','Subtraction of numbers up to 1000','End of unit assessment'] },
      { unitNo:2, unit:'Multiplication and division of whole numbers', keyCompetence:'Multiply and divide whole numbers within appropriate ranges and solve related problems.', lessons:['Multiplication tables','Multiplying 2-digit by 1-digit numbers','Division as sharing and grouping','Dividing 2-digit by 1-digit numbers','Word problems','End of unit assessment'] },
      { unitNo:3, unit:'Fractions', keyCompetence:'Identify, compare and use simple fractions in daily life situations.', lessons:['Meaning of a fraction','Unit fractions','Comparing simple fractions','Fractions of a whole and of a set','End of unit assessment'] },
      { unitNo:4, unit:'Measurement of length, mass and capacity', keyCompetence:'Measure and estimate length, mass and capacity using standard units.', lessons:['Metres and centimetres','Kilograms and grams','Litres and millilitres','Practical measuring activities','End of unit assessment'] },
      { unitNo:5, unit:'Time', keyCompetence:'Read time on the clock and use calendar information correctly.', lessons:['Hours and half hours','Quarter hours','Days, weeks and months','Using a calendar','End of unit assessment'] },
      { unitNo:6, unit:'Geometry – shapes and solids', keyCompetence:'Identify, describe and classify 2D shapes and simple 3D solids.', lessons:['2D shapes (triangle, square, rectangle, circle)','Properties of shapes','3D solids (cube, cuboid, sphere, cylinder)','Drawing and sorting shapes','End of unit assessment'] },
      { unitNo:7, unit:'Data handling', keyCompetence:'Collect, organise and interpret simple data using tables and pictographs.', lessons:['Collecting data','Tally charts','Pictographs','Interpreting simple graphs','End of unit assessment'] }
    ],
    'P5': [
      { unitNo:1, unit:'Whole numbers up to 1,000,000', keyCompetence:'Read, write, compare and perform operations on whole numbers up to one million.', lessons:['Reading and writing large numbers','Place value up to millions','Comparing and ordering','Addition and subtraction','Multiplication and division','Word problems','End of unit assessment'] },
      { unitNo:2, unit:'Factors, multiples and number properties', keyCompetence:'Find factors and multiples and use divisibility rules.', lessons:['Factors and multiples','Prime and composite numbers','Divisibility by 2, 3, 5, 10','Lowest common multiple and highest common factor','End of unit assessment'] },
      { unitNo:3, unit:'Fractions and mixed numbers', keyCompetence:'Add, subtract, multiply and divide fractions and mixed numbers and solve problems.', lessons:['Proper, improper and mixed numbers','Equivalent fractions','Addition and subtraction of fractions','Multiplication of fractions','Division of fractions','End of unit assessment'] },
      { unitNo:4, unit:'Decimal numbers', keyCompetence:'Read, write, compare and operate on decimal numbers.', lessons:['Place value of decimals','Comparing decimals','Addition and subtraction of decimals','Multiplication and division by 10, 100, 1000','End of unit assessment'] },
      { unitNo:5, unit:'Percentages', keyCompetence:'Convert between fractions, decimals and percentages and solve simple percentage problems.', lessons:['Meaning of percent','Converting fractions to percentages','Finding a percentage of a quantity','Simple percentage increase and decrease','End of unit assessment'] },
      { unitNo:6, unit:'Measurement – length, area, volume, mass, capacity', keyCompetence:'Calculate perimeter, area and volume and convert units of measurement.', lessons:['Perimeter of rectangles and squares','Area of rectangles and squares','Volume of cuboids','Units of mass and capacity','End of unit assessment'] },
      { unitNo:7, unit:'Geometry – angles, triangles and circles', keyCompetence:'Measure and classify angles and describe properties of triangles and circles.', lessons:['Types of angles','Measuring angles with a protractor','Properties of triangles','Parts of a circle','End of unit assessment'] },
      { unitNo:8, unit:'Data handling and probability', keyCompetence:'Represent and interpret data and describe simple chance events.', lessons:['Bar graphs and line graphs','Mean, mode and median (introduction)','Simple probability','End of unit assessment'] }
    ],
    'P6': [
      { unitNo:1, unit:'Whole numbers and integers', keyCompetence:'Perform operations on whole numbers and integers and solve multi-step problems.', lessons:['Large numbers and place value','Operations on whole numbers','Introduction to integers','Addition and subtraction of integers','Word problems','End of unit assessment'] },
      { unitNo:2, unit:'Fractions, decimals and percentages', keyCompetence:'Convert between and operate on fractions, decimals and percentages in real-life contexts.', lessons:['Operations on fractions','Operations on decimals','Percentages in daily life','Profit, loss and discount (simple)','End of unit assessment'] },
      { unitNo:3, unit:'Ratio, proportion and rates', keyCompetence:'Use ratio and proportion to solve problems involving sharing, scaling and rates.', lessons:['Meaning of ratio','Simplifying ratios','Direct proportion','Speed, distance and time','End of unit assessment'] },
      { unitNo:4, unit:'Algebra – simple expressions and equations', keyCompetence:'Form and simplify simple algebraic expressions and solve linear equations.', lessons:['Using letters for numbers','Simplifying expressions','Solving simple equations','Word problems leading to equations','End of unit assessment'] },
      { unitNo:5, unit:'Measurement and geometry', keyCompetence:'Calculate perimeter, area and volume of common shapes and solids and use geometric properties.', lessons:['Perimeter and area of composite shapes','Volume and capacity','Angles in triangles and quadrilaterals','Symmetry and transformations','End of unit assessment'] },
      { unitNo:6, unit:'Data handling and probability', keyCompetence:'Collect, represent and interpret data and calculate simple probabilities.', lessons:['Frequency tables','Bar, line and pie charts','Mean, median, mode and range','Simple probability experiments','End of unit assessment'] }
    ],
},
  'Social Studies': {
    'P4': [
      { unitNo:1, unit:'Socio-economic activities in our district', keyCompetence:'Compare the socio-economic activities of his/her district with those of the neighbouring districts and recognize their importance in the development of the district.', lessons:['Identification of our district and its location on the map of Rwanda','Socio-economic activities in our district','Comparing with neighbouring districts','Importance of socio-economic activities','End of unit assessment'] },
      { unitNo:2, unit:'Peace education and human rights', keyCompetence:'Explain basic human and children\u2019s rights and fight for them.', lessons:['Basic human rights','Children\u2019s rights','Promoting peace','Fighting for rights','End of unit assessment'] },
      { unitNo:4, unit:'Ten commandments of God', keyCompetence:'To adopt a culture of respecting the Ten Commandments of God.', lessons:['The Ten Commandments','Living according to the commandments','End of unit assessment'] },
      { unitNo:5, unit:'Civics and governance', keyCompetence:'Explain basic ideas of citizenship, leaders and good governance.', lessons:['Citizens and leaders','Good governance','Rights and responsibilities of citizens','End of unit assessment'] },
      { unitNo:6, unit:'Weather, flora and fauna', keyCompetence:'Describe weather, plants and animals in the environment and show care for them.', lessons:['Weather elements','Flora in our environment','Fauna in our environment','Food chains and ecosystems','End of unit assessment'] },
      { unitNo:7, unit:'Introduction to Christian and Islamic faith', keyCompetence:'A learner will be able to compare the Christian and Islamic faith.', lessons:['Basic beliefs in Christianity','Basic beliefs in Islam','Comparing Christian and Islamic faith','End of unit assessment'] },
      { unitNo:8, unit:'Infrastructure', keyCompetence:'Identify types of infrastructure and their importance in development.', lessons:['Types of infrastructure','Importance of infrastructure','Caring for public infrastructure','End of unit assessment'] },
      { unitNo:9, unit:'Traditional Rwanda', keyCompetence:'Explain social, political and economic organization in pre-colonial Rwanda.', lessons:['Social organization in traditional Rwanda','Political organization','Economic organization','End of unit assessment'] }
    ],

    'P1': [
      { unitNo:1, unit:'My family members', lessons:['Members of the nuclear family','Roles of family members','Showing respect in the family','End of unit assessment'] },
      { unitNo:2, unit:'Local government organization', lessons:['Village leaders','Duties of village leaders','Respect for leaders','End of unit assessment'] },
      { unitNo:3, unit:"God\u2019s creation", lessons:['How God created the world','When God created the world','Importance of all creatures','End of unit assessment'] },
      { unitNo:4, unit:'National symbols', lessons:['Flag of Rwanda','National anthem','Other national symbols','Respect for national symbols','End of unit assessment'] }
    ],
    'P2': [
      { unitNo:1, unit:'An extended family', lessons:['Members of an extended family up to grandparents','Roles in the extended family','Living in harmony','End of unit assessment'] },
      { unitNo:2, unit:'Local government organization', lessons:['The cell and its leaders','Duties of cell leaders','Caring for cell property','End of unit assessment'] },
      { unitNo:3, unit:'Creation of humans', lessons:['Difference between human beings and other creatures','Respect for creation','End of unit assessment'] },
      { unitNo:6, unit:'Economy', lessons:['Resources in everyday life','Not misusing resources','Saving and sharing','End of unit assessment'] }
    ],
    'P3': [
      { unitNo:1, unit:'An extended family', lessons:['Duties in the extended family','Importance of fulfilling duties','Dangers of not fulfilling duties','End of unit assessment'] },
      { unitNo:2, unit:'Our sector', lessons:['Leaders at sector level','Duties of sector leaders','Services at sector level','End of unit assessment'] },
      { unitNo:3, unit:"Respect of God\u2019s creatures", lessons:['Creatures in the environment','Caring for creatures','Respect for all creation','End of unit assessment'] }
    ],
    'P5': [
      { unitNo:1, unit:'Socio-economic activities in our province', keyCompetence:'Describe the main socio-economic activities of his/her province and explain their contribution to national development.', lessons:['Location of our province on the map of Rwanda','Main economic activities in the province','Social activities and services','Comparing provinces','Contribution to national development','End of unit assessment'] },
      { unitNo:2, unit:'Peace, conflict resolution and human rights', keyCompetence:'Demonstrate peaceful ways of resolving conflicts and respect for human rights in daily life.', lessons:['Types of conflicts in the community','Causes of conflict','Peaceful conflict resolution methods','Human and children rights','Duties that go with rights','End of unit assessment'] },
      { unitNo:3, unit:'Civics, citizenship and good governance', keyCompetence:'Explain the roles of citizens and leaders and practise good citizenship.', lessons:['Meaning of citizenship','Qualities of a good citizen','Leaders at different levels','Good governance principles','Participation in community work','End of unit assessment'] },
      { unitNo:4, unit:'Environment, weather and climate', keyCompetence:'Describe weather and climate of the region and show care for the environment.', lessons:['Weather elements and instruments','Climate of Rwanda','Effects of weather on human activities','Environmental protection','Natural disasters and safety','End of unit assessment'] },
      { unitNo:5, unit:'History of Rwanda before independence', keyCompetence:'Describe the main events and organisation of Rwanda before independence.', lessons:['Traditional social organisation','Traditional political organisation','Economic activities in pre-colonial Rwanda','Colonial period overview','End of unit assessment'] },
      { unitNo:6, unit:'Infrastructure and development', keyCompetence:'Identify major infrastructure in the province and explain its importance for development.', lessons:['Types of infrastructure','Transport and communication','Health and education facilities','Importance for development','Caring for public property','End of unit assessment'] },
      { unitNo:7, unit:'Religious and moral education', keyCompetence:'Appreciate moral values from Christian and Islamic teachings and apply them in daily life.', lessons:['Selected Bible teachings on love and honesty','Selected Islamic teachings on peace and honesty','Living moral values','Respect for different faiths','End of unit assessment'] }
    ],
    'P6': [
      { unitNo:1, unit:'Socio-economic activities in Rwanda and the region', keyCompetence:'Analyse the main socio-economic activities of Rwanda and neighbouring countries and their contribution to development.', lessons:['Map of Rwanda and East Africa','Agriculture, industry and services in Rwanda','Trade with neighbouring countries','Tourism and natural resources','Challenges and opportunities','End of unit assessment'] },
      { unitNo:2, unit:'Peace education, genocide and reconciliation', keyCompetence:'Explain the causes and consequences of the 1994 Genocide against the Tutsi and promote unity and reconciliation.', lessons:['Causes of the Genocide against the Tutsi','Consequences of the Genocide','Unity and reconciliation institutions','Role of a citizen in promoting peace','Never again message','End of unit assessment'] },
      { unitNo:3, unit:'Civics, democracy and governance in Rwanda', keyCompetence:'Describe the system of governance in Rwanda and practise democratic values.', lessons:['The Constitution of Rwanda','Organs of government','Local government structure','Elections and participation','Rights and responsibilities of citizens','End of unit assessment'] },
      { unitNo:4, unit:'Population, environment and sustainable development', keyCompetence:'Explain the relationship between population, environment and sustainable development in Rwanda.', lessons:['Population distribution and density','Environmental challenges in Rwanda','Sustainable development goals','Personal and community actions','End of unit assessment'] },
      { unitNo:5, unit:'History of Rwanda from independence to present', keyCompetence:'Describe the main political and socio-economic developments in Rwanda since independence.', lessons:['Independence and early years','The 1994 Genocide against the Tutsi','Post-genocide reconstruction','Vision 2050 and national development','End of unit assessment'] },
      { unitNo:6, unit:'Regional integration and international relations', keyCompetence:'Explain the importance of regional integration and Rwanda\'s relations with other countries.', lessons:['East African Community','African Union','International organisations','Benefits of cooperation','End of unit assessment'] },
      { unitNo:7, unit:'Moral and religious values for citizenship', keyCompetence:'Apply moral and religious values that support good citizenship and social cohesion.', lessons:['Values of honesty, respect and hard work','Teachings on peace and forgiveness','Respect for diversity','Living as a responsible citizen','End of unit assessment'] }
    ],
},
  'Religious Education': {
    'P1': [
      { unitNo:3, unit:"God\u2019s creation", lessons:['How and when God created the world','Importance of creatures','End of unit assessment'] }
    ],
    'P3': [
      { unitNo:3, unit:"Respect of God\u2019s creatures", lessons:['Respect and care for creatures','Practical activities','End of unit assessment'] }
    ]
  }
};

/* REB CBC-aligned units & lessons (from REB syllabi / schemes of work / content distribution) */
var UNITS_BY_SUBJECT = {
  'SET': [
    { unit: 'Materials and tools used at home and at school', lessons: ['Natural and artificial materials','Tools used at home','Tools used at school','Safe use and care of tools'] },
    { unit: 'Making toys and teaching and learning materials', lessons: ['Making a toy from paper','Making toys from local materials','Making simple teaching aids','Practical activities and assessment'] },
    { unit: 'Computer my friend', lessons: ['Parts of a laptop computer','Parts of a desktop computer','Uses of computer parts','Taking care of a computer','Safe sitting posture'] },
    { unit: 'Plants', lessons: ['Main parts of a plant','Importance of plant parts','Grouping plants by use','Caring for plants'] },
    { unit: 'Carpentry tools', lessons: ['Identification of carpentry tools','Use and maintenance of carpentry tools','Dangers and safety measures','End of unit assessment'] },
    { unit: 'Masonry tools', lessons: ['Identification of masonry tools','Usage and maintenance of masonry tools','Dangers of masonry tools','Practical activities'] },
    { unit: 'Light', lessons: ['How light travels','Reflection of light','Bending of light','Simple experiments on light'] },
    { unit: 'Human digestive system', lessons: ['Parts and functions of digestive system','Stages of digestion','Hygiene and healthy digestion','Digestive disorders and prevention'] },
    { unit: 'Human respiratory system', lessons: ['Parts of the respiratory system','Breathing mechanism','Keeping the respiratory system healthy'] },
    { unit: 'Matter and energy', lessons: ['States of matter','Heat and temperature','Simple forms of energy','Sources of energy in daily life'] },
    { unit: 'Programming for children', lessons: ['Introduction to Turtle Art','Drawing shapes with Scratch','Simple robotics activities','Creating a digital story'] }
  ],
  'Mathematics': [
    { unit: 'Numbers from 1 to 5', lessons: ['Counting numbers 1 to 5','Reading and writing numbers 1 to 5','Ordering and comparing','Addition and subtraction within 5'] },
    { unit: 'Numbers from 1 to 9', lessons: ['Counting and writing 1 to 9','Ordering and comparing','Addition and subtraction within 9','Word problems'] },
    { unit: 'Mathematical operations on whole numbers up to 100,000', lessons: ['Reading and writing large numbers','Place value','Addition and subtraction','Multiplication and division','Word problems'] },
    { unit: 'Positive and negative integers', lessons: ['Concept of integers','Ordering integers','Addition of integers','Subtraction of integers'] },
    { unit: 'Classifying numbers by their properties', lessons: ['Even and odd numbers','Prime and composite numbers','Factors and multiples','Divisibility rules'] },
    { unit: 'Fractions with the same denominator', lessons: ['Proper and improper fractions','Equivalent fractions','Adding fractions','Subtracting fractions'] },
    { unit: 'Decimal numbers', lessons: ['Reading and writing decimals','Comparing decimals','Adding and subtracting decimals','Decimals in real life'] },
    { unit: 'Length measurements', lessons: ['Units of length','Measuring length','Perimeter of shapes','Word problems on length'] },
    { unit: 'Sets', lessons: ['Concept of a set','Types of sets','Venn diagrams','Relations and applications'] },
    { unit: 'Sets of numbers', lessons: ['Natural and whole numbers','Integers','Rational numbers','Properties of operations'] },
    { unit: 'Geometry and measurement', lessons: ['Types of angles','Perimeter and area','Volume of simple solids','Practical measurement'] },
    { unit: 'Statistics and probability', lessons: ['Collecting data','Representing data','Mean, mode and median','Simple probability'] }
  ],
  'Physics': [
    { unit: 'Introduction to physics and measurements', lessons: ['Nature of physics','Fundamental and derived quantities','SI units','Measuring length, mass and time'] },
    { unit: 'Uniform rectilinear motion', lessons: ['Distance and displacement','Speed and velocity','Acceleration','Graphs of motion'] },
    { unit: 'Force', lessons: ['Types of forces','Effects of forces','Resultant force','Applications in daily life'] },
    { unit: "Newton's laws of motion", lessons: ["Newton's first law",'Inertia and mass',"Newton's second law","Newton's third law"] },
    { unit: 'Work, power and energy', lessons: ['Work done by a force','Power','Forms of energy','Conservation of energy'] },
    { unit: 'Simple machines', lessons: ['Levers','Pulleys','Inclined plane','Mechanical advantage'] },
    { unit: 'Heat and temperature', lessons: ['Temperature scales','Heat transfer','Expansion','Applications'] },
    { unit: 'Magnetism', lessons: ['Magnetic materials','Magnetic field','Earth as a magnet','Uses of magnets'] },
    { unit: 'Current electricity', lessons: ['Simple electric circuits','Ohm\'s law','Series and parallel','Effects of current'] },
    { unit: 'Reflection and refraction of light', lessons: ['Laws of reflection','Images in plane mirrors','Refraction through lenses','Properties of images by thin lenses'] }
  ],
  'Chemistry': [
    { unit: 'Introduction to chemistry', lessons: ['Nature of chemistry','Laboratory apparatus','Safety rules','Scientific method'] },
    { unit: 'Atomic structure', lessons: ['Structure of the atom','Electron configuration','Isotopes','Atomic number and mass number'] },
    { unit: 'Periodic table', lessons: ['Arrangement of elements','Groups and periods','Properties of groups','Trends in the periodic table'] },
    { unit: 'Chemical bonding', lessons: ['Ionic bonding','Covalent bonding','Metallic bonding','Properties of compounds'] },
    { unit: 'Acids, bases and salts', lessons: ['Properties of acids and bases','pH scale','Neutralisation','Preparation of salts'] },
    { unit: 'States of matter', lessons: ['Particle theory','Changes of state','Diffusion','Applications'] },
    { unit: 'Separation of mixtures', lessons: ['Filtration and evaporation','Distillation','Chromatography','Fractional distillation'] },
    { unit: 'Air and combustion', lessons: ['Composition of air','Combustion','Rusting','Prevention of rusting'] }
  ],
  'Biology': [
    { unit: 'Cell structure', lessons: ['Animal and plant cells','Cell organelles and functions','Differences between plant and animal cells','Cell division'] },
    { unit: 'Human digestive system', lessons: ['Parts and functions of digestive system','Enzymes in digestion','Balanced diet','Digestive disorders'] },
    { unit: 'Human respiratory system', lessons: ['Parts of the respiratory system','Mechanism of breathing','Gaseous exchange','Respiratory diseases'] },
    { unit: 'Classification of living things', lessons: ['Characteristics of living things','Kingdoms of life','Vertebrates and invertebrates','Identification keys'] },
    { unit: 'Ecology and environment', lessons: ['Ecosystems','Food chains and webs','Human impact on environment','Conservation'] },
    { unit: 'Reproduction in plants and animals', lessons: ['Asexual reproduction','Sexual reproduction in flowering plants','Human reproductive system','Responsible behaviour'] },
    { unit: 'Transport in plants and animals', lessons: ['Xylem and phloem','Transpiration','Blood circulatory system','Importance of transport'] }
  ],
  'English': [
    { unit: 'Reading comprehension', lessons: ['Identifying main ideas','Inferring meaning','Summarising a text','Answering comprehension questions'] },
    { unit: 'Grammar and writing', lessons: ['Tenses in context','Writing a formal letter','Descriptive paragraph','Narrative writing'] },
    { unit: 'Listening and speaking', lessons: ['Listening for specific information','Oral presentations','Dialogues and role play','Pronunciation practice'] },
    { unit: 'Vocabulary development', lessons: ['Word families','Synonyms and antonyms','Using context clues','Subject-specific vocabulary'] },
    { unit: 'Literature and poetry', lessons: ['Elements of a story','Understanding poems','Characters and themes','Creative response'] }
  ],
  'Kinyarwanda': [
    { unit: 'Gusoma no gusobanura', lessons: ['Gusoma inyandiko','Gushaka ibisobanuro','Guhuza n\'ubuzima busanzwe','Gusubiza ibibazo'] },
    { unit: 'Imvugo n\'inyandiko', lessons: ['Gukoresha amagambo neza','Kwandika inyandiko ngufi','Inyandiko y\'ubutumwa','Gukosora amakosa'] },
    { unit: 'Umuco n\'amateka', lessons: ['Imigenzo y\'Abanyarwanda','Indirimbo n\'imbyino','Amateka y\'igihugu','Gukunda umuco'] }
  ],
  'French': [
    { unit: 'Salutations et présentation', lessons: ['Se présenter','Saluer et prendre congé','Demander et donner des informations','Dialogues simples'] },
    { unit: 'La famille et la maison', lessons: ['Membres de la famille','La maison et les pièces','Décrire une personne','Vocabulaire de la vie quotidienne'] },
    { unit: 'Grammaire de base', lessons: ['Articles et noms','Verbes au présent','Adjectifs','Phrases simples'] }
  ],
  'Social Studies': [
    { unit: 'Our family and community', lessons: ['Members of the family','Roles in the family','Our community helpers','Living together in peace'] },
    { unit: 'Our environment', lessons: ['Natural and man-made features','Caring for the environment','Weather and seasons','Resources in our locality'] },
    { unit: 'Rwanda our country', lessons: ['Location of Rwanda','Provinces and districts','National symbols','Unity and reconciliation'] },
    { unit: 'Map work and directions', lessons: ['Map symbols and keys','Cardinal points','Measuring distance on a map','Simple sketch maps'] }
  ],
  'Religious Education': [
    { unit: 'Creation and the Creator', lessons: ['Stories of creation','Respect for creation','Human responsibility','Values of care'] },
    { unit: 'Moral values and good behaviour', lessons: ['Honesty and truthfulness','Respect for others','Helping the needy','Living in peace'] }
  ],
  'History': [
    { unit: 'Pre-colonial Rwanda', lessons: ['Social organisation','Political organisation','Cultural practices','Economic activities'] },
    { unit: 'Colonial period in Rwanda', lessons: ['Arrival of colonialists','Colonial administration','Impact on Rwandan society','Resistance'] },
    { unit: 'Post-independence Rwanda', lessons: ['Independence','Challenges after independence','Unity and reconciliation','Vision 2050'] },
    { unit: 'History and Citizenship', lessons: ['Rights and responsibilities','National identity','Peace education','Active citizenship'] }
  ],
  'Geography': [
    { unit: 'Map work', lessons: ['Map symbols and keys','Measuring distance on a map','Relief representation','Direction and bearing'] },
    { unit: 'Physical geography of Rwanda', lessons: ['Relief and drainage','Climate','Soils and vegetation','Natural resources'] },
    { unit: 'Human geography', lessons: ['Population distribution','Settlement patterns','Economic activities','Urbanisation'] },
    { unit: 'Environment and sustainability', lessons: ['Environmental problems','Conservation measures','Climate change','Sustainable development'] }
  ],
  'ICT': [
    { unit: 'Computer basics', lessons: ['Parts of a computer','Input and output devices','Storage devices','Safe use of computers'] },
    { unit: 'Word processing', lessons: ['Creating a document','Formatting text','Inserting images and tables','Saving and printing'] },
    { unit: 'Internet and online safety', lessons: ['Using a web browser','Searching for information','Email basics','Online safety and cyberbullying'] },
    { unit: 'Presentation software', lessons: ['Creating slides','Adding text and images','Transitions and animations','Presenting to an audience'] }
  ],
  'Entrepreneurship': [
    { unit: 'Business ideas', lessons: ['Identifying business opportunities','SWOT analysis','Simple business plan','Resources needed'] },
    { unit: 'Starting a small business', lessons: ['Forms of business ownership','Legal requirements','Marketing basics','Record keeping'] },
    { unit: 'Financial literacy', lessons: ['Saving and budgeting','Sources of finance','Profit and loss','Simple bookkeeping'] }
  ],
  'Creative Arts': [
    { unit: 'Drawing and painting', lessons: ['Basic drawing techniques','Use of colours','Still life drawing','Creative composition'] },
    { unit: 'Crafts and local materials', lessons: ['Weaving and basketry','Clay modelling','Recycling materials','Exhibition of work'] },
    { unit: 'Music, dance and drama', lessons: ['Rwandan traditional songs','Basic dance movements','Simple drama performance','Teamwork in performing arts'] }
  ],
  'Physical Education': [
    { unit: 'Basic motor skills', lessons: ['Running and jumping','Throwing and catching','Balancing activities','Games for coordination'] },
    { unit: 'Ball games', lessons: ['Football basic skills','Netball / basketball basics','Teamwork and fair play','Simple rules of the game'] },
    { unit: 'Health and fitness', lessons: ['Warm-up and cool-down','Benefits of exercise','Personal hygiene in sports','Safety during play'] }
  ],
  'Kiswahili': [
    { unit: 'Salamu na kujitambulisha', lessons: ['Salamu za asubuhi na jioni','Kujitambulisha','Kuuliza na kujibu maswali','Mazungumzo rahisi'] },
    { unit: 'Familia na nyumbani', lessons: ['Wanachama wa familia','Sehemu za nyumba','Maelezo ya watu','Msamiati wa kila siku'] },
    { unit: 'Kusoma na kuandika', lessons: ['Kusoma sentensi','Kuandika insha fupi','Sarufi ya msingi','Kurekebisha makosa'] }
  ],
  'Economics': [
    { unit: 'Introduction to economics', lessons: ['What is economics','Needs and wants','Scarcity and choice','Opportunity cost'] },
    { unit: 'Production', lessons: ['Factors of production','Types of production','Division of labour','Productivity'] },
    { unit: 'Market and price', lessons: ['Demand and supply','Price determination','Market structures','Consumer behaviour'] }
  ],
  'General Paper': [
    { unit: 'Critical thinking and current affairs', lessons: ['Analysing news articles','Forming reasoned opinions','Debate skills','Writing an argumentative essay'] },
    { unit: 'Society and development', lessons: ['Education and development','Health issues','Gender and equity','Globalisation'] }
  ]
};

function fillSelect(sel, items, placeholder) {
  if (!sel) return;
  var keep = sel.value;
  sel.innerHTML = '<option value="">' + (placeholder || 'Select') + '</option>';
  (items || []).forEach(function (it) {
    var o = document.createElement('option');
    o.value = it;
    o.textContent = it;
    sel.appendChild(o);
  });
  if (keep && items && items.indexOf(keep) >= 0) sel.value = keep;
}

function onLevelChange() {
  var level = document.getElementById('level').value;
  fillSelect(document.getElementById('class'), CLASS_BY_LEVEL[level] || [], 'Select class (P1–P6 / S1–S6)');
  fillSelect(document.getElementById('subject'), SUBJECTS_BY_LEVEL[level] || [], 'Select subject');
  document.getElementById('unitTitleSelect').innerHTML = '<option value="">Select unit title from REB syllabus (or type below)</option>';
  document.getElementById('lessonTitleSelect').innerHTML = '<option value="">Select lesson title from unit (or type below)</option>';
  var ut = document.getElementById('unitTitle');
  var lt = document.getElementById('lessonTitle');
  if (ut) ut.value = '';
  if (lt) lt.value = '';
}

function onClassChange() {
  // Class change refreshes REB units/lessons for the selected subject
  if (getSubjectValue()) onSubjectChange();
}

function getUnitsForSelection() {
  var subj = getSubjectValue();
  var cls = (document.getElementById('class') && document.getElementById('class').value) || '';
  var level = (document.getElementById('level') && document.getElementById('level').value) || '';
  // Aliases so form subject names match REB banks
  var aliases = {
    'Creative Arts': 'Music',
    'Creative Arts / Music': 'Music',
    'Music': 'Music',
    'Physical Education and Sports': 'Physical Education',
    'PES': 'Physical Education',
    'Social Studies and Religious Education': 'Social Studies',
    'Social and Religious Education': 'Social Studies',
    'SSRE': 'Social Studies',
    'Science and Elementary Technology': 'SET',
    'Français': 'French',
    'Ikinyarwanda': 'Kinyarwanda'
  };
  var subjKey = aliases[subj] || subj;
  var units = [];
  if (typeof UNITS_BY_SUBJECT_CLASS !== 'undefined' && UNITS_BY_SUBJECT_CLASS[subjKey] && UNITS_BY_SUBJECT_CLASS[subjKey][cls]) {
    units = UNITS_BY_SUBJECT_CLASS[subjKey][cls];
  } else if (typeof UNITS_BY_SUBJECT_CLASS !== 'undefined' && UNITS_BY_SUBJECT_CLASS[subj] && UNITS_BY_SUBJECT_CLASS[subj][cls]) {
    units = UNITS_BY_SUBJECT_CLASS[subj][cls];
  } else if (UNITS_BY_SUBJECT[subjKey]) {
    units = UNITS_BY_SUBJECT[subjKey];
  } else if (UNITS_BY_SUBJECT[subj]) {
    units = UNITS_BY_SUBJECT[subj];
  }
  if (level && units.length) {
    var filtered = units.filter(function (u) {
      if (!u.levels || !u.levels.length) return true;
      return u.levels.indexOf(level) >= 0 || u.levels.indexOf(cls) >= 0;
    });
    if (filtered.length) units = filtered;
  }
  return units;
}

function onSubjectChange() {
  var units = getUnitsForSelection();
  var us = document.getElementById('unitTitleSelect');
  us.innerHTML = '<option value="">Select unit from REB syllabus (Unit number + title)</option>';
  units.forEach(function (u, idx) {
    var num = (u.unitNo != null) ? u.unitNo : (idx + 1);
    var o = document.createElement('option');
    o.value = u.unit;
    o.setAttribute('data-unit-no', String(num));
    o.textContent = 'Unit ' + num + ': ' + u.unit;
    us.appendChild(o);
  });
  document.getElementById('lessonTitleSelect').innerHTML = '<option value="">Select lesson from unit (Lesson number + title)</option>';
  document.getElementById('unitTitle').value = '';
  document.getElementById('lessonTitle').value = '';
}

function onUnitSelect() {
  var us = document.getElementById('unitTitleSelect');
  var unit = us.value;
  if (unit) document.getElementById('unitTitle').value = unit;
  // Auto-fill Unit No from REB list order
  try {
    var opt = us.options[us.selectedIndex];
    var uno = opt && opt.getAttribute('data-unit-no');
    if (uno && document.getElementById('unitNo')) document.getElementById('unitNo').value = uno;
  } catch (e) {}
  var units = getUnitsForSelection();
  var found = units.filter(function (u) { return u.unit === unit; })[0];
  var ls = document.getElementById('lessonTitleSelect');
  ls.innerHTML = '<option value="">Select lesson from unit (Lesson number + title)</option>';
  document.getElementById('lessonTitle').value = '';
  if (found && found.lessons) {
    found.lessons.forEach(function (L, idx) {
      var o = document.createElement('option');
      o.value = L;
      o.setAttribute('data-lesson-no', String(idx + 1));
      o.textContent = 'Lesson ' + (idx + 1) + ': ' + L;
      ls.appendChild(o);
    });
    // auto total lessons in unit
    try {
      if (document.getElementById('totalLessons')) {
        document.getElementById('totalLessons').value = found.lessons.length;
      }
    } catch (e2) {}
  }
  try {
    var kc = document.getElementById('keyCompetence');
    if (kc && found && found.keyCompetence) {
      kc.value = found.keyCompetence;
      kc.placeholder = 'From REB scheme of work (you may edit)';
    }
  } catch (e3) {}
}

function onLessonSelect() {
  var ls = document.getElementById('lessonTitleSelect');
  var L = ls.value;
  if (L) document.getElementById('lessonTitle').value = L;
  try {
    var opt = ls.options[ls.selectedIndex];
    var lno = opt && opt.getAttribute('data-lesson-no');
    if (lno && document.getElementById('lessonNo')) document.getElementById('lessonNo').value = lno;
    // if total still empty, use number of options - 1 (placeholder)
    var tl = document.getElementById('totalLessons');
    if (tl && !tl.value && ls.options.length > 1) tl.value = String(ls.options.length - 1);
  } catch (e) {}
}

function getSubjectValue() {
  var c = (document.getElementById('subjectCustom') && document.getElementById('subjectCustom').value.trim()) || '';
  if (c) return c;
  return (document.getElementById('subject') && document.getElementById('subject').value) || '';
}

function onSelfEvalSelect() {
  var key = document.getElementById('selfEvalSelect').value;
  var topic = (document.getElementById('lessonTitle').value || document.getElementById('unitTitle').value || 'the lesson').trim();
  var unit = (document.getElementById('unitTitle').value || topic).trim();
  var map = {
    excellent: 'The instructional objectives on "' + topic + '" were fully achieved. Learners participated actively in exploring the unit "' + unit + '". Group work and presentations were successful. I will maintain this approach and add more extension tasks next time.',
    good: 'Most objectives related to "' + topic + '" were achieved. Learners showed good understanding of the unit "' + unit + '". A few learners still need practice. Next lesson I will give extra support to slower learners during group tasks.',
    average: 'Objectives on "' + topic + '" were partially achieved. Some learners struggled with key ideas in the unit "' + unit + '". Next time I will use more examples and allow more guided practice before independent work.',
    needs_improve: 'Achievement of objectives on "' + topic + '" was limited. I will reteach the main points of the unit "' + unit + '" using simpler steps, more visual aids and closer monitoring of group work.',
    mixed: 'Group work on "' + topic + '" was strong, but individual practice showed gaps in the unit "' + unit + '". Next lesson I will balance group tasks with clearer individual exercises and feedback.',
    time: 'Content on "' + topic + '" was relevant, but time for development was short. Next time I will protect more minutes for practice and presentations within the unit "' + unit + '".'
  };
  var txt = map[key] || '';
  document.getElementById('selfEvalInput').value = txt;
  if (txt) applySelfEvaluationSilent();
}


/** Build instructional objectives: "By using [materials], … correctly / accurately / …" */
function formatInstructionalObjectives(data, mats) {
  var lang = data.lang || 'en';
  var topic = (data.lessonTitle || data.unitTitle || 'the topic').trim();
  var cls = (data.class || '').trim();
  var materials = (mats || data.materials || '').trim();
  if (!materials) {
    materials = (data.subject || 'subject') + " learner's book, chalkboard, charts and available teaching aids";
  }
  // shorten materials for the "By using" phrase if very long
  var matShort = materials.length > 120 ? materials.slice(0, 117) + '…' : materials;

  if (lang === 'rw') {
    return '1. Bakoresheje ' + matShort + ', abanyeshuri ba ' + (cls || 'ishuri') +
      ' bazashobora kumenya no gusobanura ingingo z\'ingenzi za "' + topic + '" neza no mu buryo bwukuri.\n' +
      '2. Bakoresheje ' + matShort + ', abanyeshuri bazashobora gukoresha ibitekerezo ku "' + topic +
      '" mu mirimo y\'amatsinda kandi bakagira ikinamico neza.\n' +
      '3. Bakoresheje ' + matShort + ', abanyeshuri bazashobora gukora imyitozo ku "' + topic + '" mu buryo bukwiye.';
  }
  if (lang === 'fr') {
    return '1. En utilisant ' + matShort + ', les apprenants de ' + (cls || 'la classe') +
      ' seront capables d\'identifier et d\'expliquer correctement les points clés de « ' + topic + ' ».\n' +
      '2. En utilisant ' + matShort + ', les apprenants appliqueront avec exactitude les idées sur « ' + topic +
      ' » lors du travail de groupe et présenteront leurs résultats clairement.\n' +
      '3. En utilisant ' + matShort + ', les apprenants réaliseront des exercices sur « ' + topic + ' » de manière appropriée.';
  }
  if (lang === 'sw') {
    return '1. Kwa kutumia ' + matShort + ', wanafunzi wa ' + (cls || 'darasa') +
      ' wataweza kutambua na kueleza pointi muhimu za "' + topic + '" kwa usahihi.\n' +
      '2. Kwa kutumia ' + matShort + ', wanafunzi watatumia mawazo kuhusu "' + topic +
      '" katika kazi za vikundi na kuwasilisha matokeo kwa usahihi.\n' +
      '3. Kwa kutumia ' + matShort + ', wanafunzi watafanya mazoezi kuhusu "' + topic + '" ipasavyo.';
  }
  // English (default) — starts with "By using", ends with correctly / accurately / accordingly
  return '1. By using ' + matShort + ', ' + (cls || 'learners') +
    ' will be able to identify and explain the key points of "' + topic + '" correctly.\n' +
    '2. By using ' + matShort + ', learners will apply ideas about "' + topic +
    '" in guided practice and group work accurately.\n' +
    '3. By using ' + matShort + ', learners will complete simple tasks and present findings on "' + topic +
    '" accordingly.';
}

function buildSmartContent(data) {
  var lang = (data.lang || document.getElementById('languageSelect').value || 'en');
  if (!T[lang]) lang = 'en';
  var topic = (data.lessonTitle || data.unitTitle || 'the topic').trim();
  var unit = (data.unitTitle || topic).trim();
  var subj = (data.subject || '').trim();
  var cls = (data.class || 'learners').trim();
  var low = (topic + ' ' + unit + ' ' + subj).toLowerCase();

  var domain = 'general';
  if (/digest|nutrition|health|body|hygiene|disease|blood|cell|plant|animal|ecosystem|force|energy|light|sound|electric|magnet|matter|chemical|water|soil|weather|environment|science|set|biology|physics|chemistry|organ|system|food|mechanics|machine|tool|pollution|air|cow|goat|agricultur/.test(low)) domain = 'science';
  else if (/math|number|fraction|algebra|geometry|measure|percentage|ratio|statistic|equation|shape|volume|area|digit|angle|graph|integer|decimal|perimeter/.test(low)) domain = 'math';
  else if (/english|kinyarwanda|french|kiswahili|language|grammar|writing|reading|poem|story|comprehension|vocabulary|oral|letter|salut|parler|greet|introduc/.test(low)) domain = 'language';
  else if (/history|geography|civic|society|government|map|culture|peace|rights|citizen|continent|climate|social|genocide|governance|infrastructure|province|district/.test(low)) domain = 'social';
  else if (/art|music|sport|physical|draw|dance|drama|creative|motor|athlet|football|gymnast|rhythm|melody|song/.test(low)) domain = 'arts';
  else if (/ict|computer|software|internet|coding|digital|word process|spreadsheet|program/.test(low)) domain = 'ict';
  else if (/entrepreneur|business|account|market|money|finance/.test(low)) domain = 'business';

  var suffixes = {
    en: ['correctly', 'accurately', 'accordingly', 'with confidence', 'in an organised way'],
    rw: ['neza', 'mu buryo bukwiye', 'mu buryo bwizewe', 'n\'ubwitonzi'],
    fr: ['correctement', 'avec précision', 'de manière appropriée', 'avec confiance'],
    sw: ['kwa usahihi', 'kwa makini', 'kwa njia ifaayo', 'kwa ujasiri']
  };
  var suf = suffixes[lang] || suffixes.en;
  function pickSuf(i) { return suf[i % suf.length]; }

  var matShort = {
    en: 'the learner\'s book, teacher\'s guide, chalkboard, charts and relevant local materials',
    rw: 'igitabo cy\'umunyeshuri, icy\'umwarimu, ikibaho, amakarita n\'ibikoresho byo mu gace',
    fr: 'le livre de l\'élève, le guide de l\'enseignant, le tableau, les affiches et le matériel local',
    sw: 'kitabu cha mwanafunzi, mwongozo wa mwalimu, ubao, chati na vifaa vya mitaa'
  };
  var matsPhrase = matShort[lang] || matShort.en;

  // Richer domain-aware key competence
  var keyMap = {
    en: {
      science: 'On the lesson "' + topic + '" (unit "' + unit + '"), learners identify, explain and apply scientific ideas about "' + topic + '", use appropriate methods and materials safely, observe carefully, record findings, and show curiosity, accuracy and respect for the environment and living things.',
      math: 'On the lesson "' + topic + '" (unit "' + unit + '"), learners understand and accurately apply mathematical concepts related to "' + topic + '", solve problems using clear steps, check their answers, and show confidence, precision and perseverance in mathematics.',
      language: 'On the lesson "' + topic + '" (unit "' + unit + '"), learners use listening, speaking, reading and writing skills related to "' + topic + '", communicate ideas clearly and politely, expand vocabulary, and show respect for language and culture.',
      social: 'On the lesson "' + topic + '" (unit "' + unit + '"), learners explain and analyse ideas about "' + topic + '", relate them to family, community and national life, promote peace and responsibility, and practise good citizenship values.',
      arts: 'On the lesson "' + topic + '" (unit "' + unit + '"), learners create, perform and appreciate activities related to "' + topic + '", express ideas creatively, work safely with others, and value physical well-being, culture and the arts.',
      ict: 'On the lesson "' + topic + '" (unit "' + unit + '"), learners use digital tools related to "' + topic + '" safely and responsibly, complete simple ICT tasks with accuracy, protect personal information, and show good digital citizenship.',
      business: 'On the lesson "' + topic + '" (unit "' + unit + '"), learners apply entrepreneurial ideas about "' + topic + '", make simple and honest business decisions, work with others, and show initiative, creativity and responsibility.',
      general: 'On the lesson "' + topic + '" (unit "' + unit + '"), learners understand and apply knowledge and skills about "' + topic + '", work collaboratively, communicate their ideas, and show positive attitudes towards learning and lifelong growth.'
    },
    rw: {
      science: 'Ku isomo "' + topic + '" (igice "' + unit + '"), abanyeshuri kumenya, gusobanura no gukoresha ibitekerezo bya siyansi ku "' + topic + '", bakoresha uburyo n\'ibikoresho mu buryo bwizewe, bakurikirana neza, bandika ibyo babonye, kandi bagaragaza amatsiko, ubwitonzi n\'kubaha ibidukikije.',
      math: 'Ku isomo "' + topic + '" (igice "' + unit + '"), abanyeshuri gusobanukirwa no gukoresha neza imibare ijyanye na "' + topic + '", bakemura ibibazo mu ntambwe zisobanutse, basuzume ibisubizo, kandi bagaragaza ubwitonzi n\'ubwenge mu imibare.',
      language: 'Ku isomo "' + topic + '" (igice "' + unit + '"), abanyeshuri gukoresha ubuhanga bwo kumva, kuvuga, gusoma no kwandika ku "' + topic + '", batanga ubutumwa neza kandi mu buryo bwubaha, kongera ububiko bw\'amagambo, bakubaha ururimi n\'umuco.',
      social: 'Ku isomo "' + topic + '" (igice "' + unit + '"), abanyeshuri gusobanura no gusesengura ibitekerezo ku "' + topic + '", babihuze n\'umuryango, umudugudu n\'igihugu, bashyigikira amahoro n\'inshingano, bakora nk\'abenegihugu beza.',
      arts: 'Ku isomo "' + topic + '" (igice "' + unit + '"), abanyeshuri gukora, kwerekana no gushima ibikorwa bijyanye na "' + topic + '", bakoresheje ubuhanga, bakora mu mutekano, kandi bashyigikira ubuzima bwiza, umuco n\'ubugeni.',
      ict: 'Ku isomo "' + topic + '" (igice "' + unit + '"), abanyeshuri gukoresha ibikoresho bya mudasobwa ku "' + topic + '" mu buryo bwizewe, bakora imirimo yoroshye ya ICT neza, bakinga amakuru yabo, bagaragaza imyitwarire myiza ku ikoranabuhanga.',
      business: 'Ku isomo "' + topic + '" (igice "' + unit + '"), abanyeshuri gukoresha ibitekerezo by\'ubucuruzi ku "' + topic + '", bafata ibyemezo by\'ukuri, bakora hamwe, bagaragaza ubushobozi n\'inshingano.',
      general: 'Ku isomo "' + topic + '" (igice "' + unit + '"), abanyeshuri gusobanukirwa no gukoresha ubumenyi n\'ubuhanga ku "' + topic + '", bakora hamwe, batanga ibitekerezo, bagaragaza imyitwarire myiza yo kwiga.'
    },
    fr: {
      science: 'Dans la leçon « ' + topic + ' » (unité « ' + unit + ' »), les apprenants identifier, expliquer et appliquer des idées scientifiques sur « ' + topic + ' », utiliser des méthodes et matériels de façon sûre, observer, noter les résultats et montrer curiosité, précision et respect de l\'environnement.',
      math: 'Dans la leçon « ' + topic + ' » (unité « ' + unit + ' »), les apprenants comprendre et appliquer avec précision les notions mathématiques liées à « ' + topic + ' », résoudre des problèmes par étapes claires, vérifier leurs réponses et montrer confiance et rigueur.',
      language: 'Dans la leçon « ' + topic + ' » (unité « ' + unit + ' »), les apprenants utiliser l\'écoute, la parole, la lecture et l\'écriture autour de « ' + topic + ' », communiquer clairement et poliment, enrichir le vocabulaire et respecter la langue et la culture.',
      social: 'Dans la leçon « ' + topic + ' » (unité « ' + unit + ' »), les apprenants expliquer et analyser des idées sur « ' + topic + ' », les relier à la famille, à la communauté et à la nation, promouvoir la paix et la responsabilité, et pratiquer une bonne citoyenneté.',
      arts: 'Dans la leçon « ' + topic + ' » (unité « ' + unit + ' »), les apprenants créer, présenter et apprécier des activités liées à « ' + topic + ' », s\'exprimer avec créativité, travailler en sécurité et valoriser le bien-être, la culture et les arts.',
      ict: 'Dans la leçon « ' + topic + ' » (unité « ' + unit + ' »), les apprenants utiliser des outils numériques liés à « ' + topic + ' » de façon sûre et responsable, réaliser des tâches TIC simples avec exactitude et faire preuve de bonne citoyenneté numérique.',
      business: 'Dans la leçon « ' + topic + ' » (unité « ' + unit + ' »), les apprenants appliquer des idées entrepreneuriales sur « ' + topic + ' », prendre des décisions simples et honnêtes, collaborer et montrer initiative et responsabilité.',
      general: 'Dans la leçon « ' + topic + ' » (unité « ' + unit + ' »), les apprenants comprendre et appliquer des connaissances sur « ' + topic + ' », collaborer, communiquer leurs idées et montrer une attitude positive envers l\'apprentissage.'
    },
    sw: {
      science: 'Katika somo "' + topic + '" (kitengo "' + unit + '"), wanafunzi kutambua, kueleza na kutumia mawazo ya sayansi kuhusu "' + topic + '", kutumia mbinu na vifaa kwa usalama, kuchunguza, kurekodi matokeo na kuonyesha udadisi, usahihi na heshima kwa mazingira.',
      math: 'Katika somo "' + topic + '" (kitengo "' + unit + '"), wanafunzi kuelewa na kutumia kwa usahihi dhana za hisabati zinazohusiana na "' + topic + '", kutatua matatizo kwa hatua wazi, kukagua majibu na kuonyesha ujasiri na umakini.',
      language: 'Katika somo "' + topic + '" (kitengo "' + unit + '"), wanafunzi kutumia ustadi wa kusikiliza, kuzungumza, kusoma na kuandika kuhusu "' + topic + '", kuwasiliana kwa uwazi na adabu, kukuza msamiati na kuheshimu lugha na utamaduni.',
      social: 'Katika somo "' + topic + '" (kitengo "' + unit + '"), wanafunzi kueleza na kuchambua mawazo kuhusu "' + topic + '", kuyahusisha na familia, jamii na taifa, kukuza amani na uwajibikaji, na kuonyesha uraia bora.',
      arts: 'Katika somo "' + topic + '" (kitengo "' + unit + '"), wanafunzi kuunda, kuonyesha na kuthamini shughuli zinazohusiana na "' + topic + '", kujieleza kwa ubunifu, kufanya kazi kwa usalama na kuthamini afya, utamaduni na sanaa.',
      ict: 'Katika somo "' + topic + '" (kitengo "' + unit + '"), wanafunzi kutumia zana za kidijitali kuhusu "' + topic + '" kwa usalama na uwajibikaji, kukamilisha kazi rahisi za TEHAMA kwa usahihi na kuonyesha uraia bora wa kidijitali.',
      business: 'Katika somo "' + topic + '" (kitengo "' + unit + '"), wanafunzi kutumia mawazo ya ujasiriamali kuhusu "' + topic + '", kufanya maamuzi rahisi na ya uaminifu, kushirikiana na kuonyesha ubunifu na uwajibikaji.',
      general: 'Katika somo "' + topic + '" (kitengo "' + unit + '"), wanafunzi kuelewa na kutumia maarifa kuhusu "' + topic + '", kushirikiana, kuwasilisha mawazo yao na kuonyesha mtazamo chanya wa kujifunza.'
    }
  };

  var keys = keyMap[lang] || keyMap.en;
  var keyCompetence = keys[domain] || keys.general;
  // If a key competence was already written in the form / REB field, keep it exactly
  if (data.keyCompetence && String(data.keyCompetence).trim().length >= 20) {
    keyCompetence = String(data.keyCompetence).trim();
  }

  // Objectives: By using materials... ends with correctly/accurately...
  var objEn = '1. By using ' + matsPhrase + ', ' + cls + ' will be able to identify, describe and explain the main ideas of the lesson "' + topic + '" in unit "' + unit + '" ' + pickSuf(0) + '.\n'
    + '2. By using ' + matsPhrase + ', learners will practise and apply skills and knowledge related to "' + topic + '" through guided tasks, examples and group work ' + pickSuf(1) + '.\n'
    + '3. By using ' + matsPhrase + ', learners will present their work on "' + topic + '", answer questions and use feedback to improve ' + pickSuf(2) + '.';
  var objRw = '1. Bakoresheje ' + matsPhrase + ', abanyeshuri ba ' + cls + ' bazashobora kumenya no gusobanura ingingo z\'ingenzi za "' + topic + '" mu gice "' + unit + '" ' + pickSuf(0) + '.\n'
    + '2. Bakoresheje ' + matsPhrase + ', abanyeshuri bazakoresha ubuhanga ku "' + topic + '" binyuze mu mirimo n\'amatsinda ' + pickSuf(1) + '.\n'
    + '3. Bakoresheje ' + matsPhrase + ', abanyeshuri bazatanga ikinamico ku "' + topic + '" bakasubiza ibibazo ' + pickSuf(2) + '.';
  var objFr = '1. En utilisant ' + matsPhrase + ', les élèves de ' + cls + ' pourront identifier et expliquer les idées principales de « ' + topic + ' » dans l\'unité « ' + unit + ' » ' + pickSuf(0) + '.\n'
    + '2. En utilisant ' + matsPhrase + ', les apprenants appliqueront les compétences liées à « ' + topic + ' » par des tâches guidées et le travail de groupe ' + pickSuf(1) + '.\n'
    + '3. En utilisant ' + matsPhrase + ', les apprenants présenteront leurs résultats sur « ' + topic + ' » et répondront aux commentaires ' + pickSuf(2) + '.';
  var objSw = '1. Kwa kutumia ' + matsPhrase + ', wanafunzi wa ' + cls + ' wataweza kutambua na kueleza mawazo makuu ya "' + topic + '" katika kitengo "' + unit + '" ' + pickSuf(0) + '.\n'
    + '2. Kwa kutumia ' + matsPhrase + ', wanafunzi watatumia stadi zinazohusiana na "' + topic + '" kupitia kazi za kuongozwa na kazi za vikundi ' + pickSuf(1) + '.\n'
    + '3. Kwa kutumia ' + matsPhrase + ', wanafunzi watawasilisha matokeo yao kuhusu "' + topic + '" na kujibu maoni ' + pickSuf(2) + '.';
  var objMap = { en: objEn, rw: objRw, fr: objFr, sw: objSw };
  var objective = objMap[lang] || objEn;

  var matsEn = (subj || 'Subject') + ' Learner\'s Book and Teacher\'s Guide (REB, new curriculum); chalkboard and markers; charts, pictures or models related to "' + topic + '"; real objects or local materials linked to unit "' + unit + '"; exercise books and pens.';
  var matsRw = 'Igitabo cy\'umunyeshuri n\'icy\'umwarimu (REB, porogaramu nshya) — ' + (subj || 'Isomo') + '; ikibaho n\'amakara; amakarita, amashusho cyangwa modeli za "' + topic + '"; ibikoresho by\'ukuri byo mu gace; amakaye n\'ibikoresho byo kwandika.';
  var matsFr = 'Livre de l\'élève et guide de l\'enseignant (REB, nouveau curriculum) — ' + (subj || 'Matière') + ' ; tableau et craies/marqueurs ; affiches, images ou modèles sur « ' + topic + ' » ; objets réels ou matériel local ; cahiers et stylos.';
  var matsSw = 'Kitabu cha mwanafunzi na mwongozo wa mwalimu (REB, mtaala mpya) — ' + (subj || 'Somo') + '; ubao na kalamu; chati, picha au mifano kuhusu "' + topic + '"; vifaa vya mitaa; daftari na kalamu.';
  var mats = ({ en: matsEn, rw: matsRw, fr: matsFr, sw: matsSw })[lang] || matsEn;

  var refsEn = 'Rwanda Education Board (REB). Competence-Based Curriculum syllabus and scheme of work — ' + (subj || 'Subject') + ', ' + cls + '.\nRwanda Education Board. Learner\'s Book and Teacher\'s Guide — ' + (subj || 'Subject') + ' (Nouvelle Édition / new version). Kigali: REB.';
  var refsRw = 'Ikigo cy\'Integanyanyigisho mu Rwanda (REB). Porogaramu ishingiye ku bushobozi n\'inganzo y\'isomo — ' + (subj || 'Isomo') + ', ' + cls + '.\nREB. Igitabo cy\'umunyeshuri n\'icy\'umwarimu — ' + (subj || 'Isomo') + '. Kigali: REB.';
  var refsFr = 'Conseil de l\'Éducation de base du Rwanda (REB). Programme basé sur les compétences et répartition des contenus — ' + (subj || 'Matière') + ', ' + cls + '.\nREB. Livre de l\'élève et guide de l\'enseignant — ' + (subj || 'Matière') + '. Kigali: REB.';
  var refsSw = 'Bodi ya Elimu ya Msingi Rwanda (REB). Mtaala wa uwezo na mpango wa kazi — ' + (subj || 'Somo') + ', ' + cls + '.\nREB. Kitabu cha mwanafunzi na mwongozo wa mwalimu — ' + (subj || 'Somo') + '. Kigali: REB.';
  var references = ({ en: refsEn, rw: refsRw, fr: refsFr, sw: refsSw })[lang] || refsEn;

  var activityNote = ({
    en: 'Active learning on "' + topic + '": clear explanation, demonstration, pair/group practice, presentation and constructive feedback so every learner participates.',
    rw: 'Kwiga bikora ku "' + topic + '": gusobanura, kwerekana, gukora mu matsinda, ikinamico n\'inyunganizi kugira ngo umunyeshuri wese agire uruhare.',
    fr: 'Apprentissage actif sur « ' + topic + ' » : explication, démonstration, travail en binômes/groupes, présentation et feedback pour la participation de tous.',
    sw: 'Kujifunza kwa vitendo kuhusu "' + topic + '": maelezo, onyesho, kazi za vikundi, uwasilishaji na maoni ili kila mwanafunzi ashiriki.'
  })[lang];

  // Richer phased activities (topic-aware, enough lines to fill A4 when scaled)
  var introT = ({
    en: [
      'Greet learners warmly, check attendance and ensure the classroom is ready for learning.',
      'State the lesson title "' + topic + '" and link it clearly to unit "' + unit + '" and subject ' + (subj || '') + '.',
      'Ask short oral questions to recall prior knowledge connected to "' + topic + '".',
      'Share the instructional objectives and explain what success looks like in this lesson.',
      'Motivate learners by showing a picture, object or short example related to "' + topic + '".'
    ],
    rw: [
      'Bashimiye abanyeshuri, reba ko bahari kandi class iteguye.',
      'Vuga umutwe w\'isomo "' + topic + '" uhuze n\'igice "' + unit + '".',
      'Baza ibibazo by\'igufi byo kwibuka ibyiganirijwe mbere bijyanye na "' + topic + '".',
      'Sobanura intego z\'isomo n\'icyo bategerejwe kugeraho.',
      'Bahugure abanyeshuri hakoreshejwe ishusho cyangwa urugero rwa "' + topic + '".'
    ],
    fr: [
      'Saluer les élèves, vérifier la présence et préparer la classe.',
      'Annoncer le titre « ' + topic + ' » et le relier à l\'unité « ' + unit + ' ».',
      'Poser des questions orales de rappel sur les acquis liés à « ' + topic + ' ».',
      'Présenter les objectifs d\'apprentissage et les critères de réussite.',
      'Motiver avec une image, un objet ou un exemple concret sur « ' + topic + ' ».'
    ],
    sw: [
      'Salimu wanafunzi, kagua mahudhurio na andaa darasa.',
      'Taja kichwa cha somo "' + topic + '" na kiunganishe na kitengo "' + unit + '".',
      'Uliza maswali mafupi ya kukumbusha maarifa ya awali kuhusu "' + topic + '".',
      'Eleza malengo ya somo na kinachotarajiwa.',
      'Hamasisisha kwa picha, kitu halisi au mfano kuhusu "' + topic + '".'
    ]
  })[lang];

  var introL = ({
    en: [
      'Respond to the greeting, settle down and show readiness to learn.',
      'Listen carefully to the lesson title, unit link and objectives.',
      'Answer recall questions and share what they already know about "' + topic + '".',
      'Ask clarifying questions if the objectives are not clear.',
      'Observe the motivating example and connect it to their experience.'
    ],
    rw: [
      'Basubize ikaze, bicare neza bakitegura kwiga.',
      'Tega amatwi ku mutwe w\'isomo, igice n\'intego.',
      'Subiza ibibazo byo kwibuka, bagire icyo bazi ku "' + topic + '".',
      'Babaze niba intego zidasobanutse.',
      'Reba urugero rw\'inkunga bakihuza n\'ubunararibonye.'
    ],
    fr: [
      'Répondre à la salutation et se préparer à apprendre.',
      'Écouter le titre, le lien avec l\'unité et les objectifs.',
      'Répondre aux questions de rappel et partager leurs connaissances sur « ' + topic + ' ».',
      'Demander des précisions si nécessaire.',
      'Observer l\'exemple motivant et le relier à leur vécu.'
    ],
    sw: [
      'Jibu salamu na ujiandae kujifunza.',
      'Sikiliza kichwa cha somo, kitengo na malengo.',
      'Jibu maswali ya ukumbusho na eleza wanachojua kuhusu "' + topic + '".',
      'Uliza maswali ya ufafanuzi inapohitajika.',
      'Angalia mfano wa motisha na uhusishe na uzoefu wao.'
    ]
  })[lang];

  var introC = ({
    en: [
      'Communication: clear oral interaction between teacher and learners about "' + topic + '".',
      'Critical thinking: recalling and linking prior knowledge to the new lesson.',
      'Inclusive education: every learner is invited to speak and is listened to with respect.'
    ],
    rw: [
      'Itumanaho: ikiganiro gisobanutse hagati y\'umwarimu n\'abanyeshuri ku "' + topic + '".',
      'Gtekereza cyimbitse: kwibuka no guhuza ubumenyi bwa mbere n\'isomo rishya.',
      'Uburezi bw\'abantu bose: umunyeshuri wese ahaurwa kugira icyo avuga.'
    ],
    fr: [
      'Communication : échange oral clair sur « ' + topic + ' ».',
      'Pensée critique : rappel et lien avec les acquis antérieurs.',
      'Éducation inclusive : chaque apprenant est invité à participer.'
    ],
    sw: [
      'Mawasiliano: majadiliano wazi kuhusu "' + topic + '".',
      'Fikra kritikal: kukumbuka na kuunganisha maarifa ya awali.',
      'Elimu jumuishi: kila mwanafunzi anahimizwa kushiriki.'
    ]
  })[lang];

  var devT = ({
    en: [
      'Explain the main concepts of "' + topic + '" using simple language, examples and teaching aids from the REB materials.',
      'Demonstrate key steps or illustrate ideas on the board so learners see how "' + topic + '" works in practice.',
      'Organise learners in pairs or small groups to complete a guided task or discussion focused on "' + topic + '".',
      'Move around the class to support groups, ask probing questions and correct misconceptions early.',
      'Invite groups to present short findings and give specific, constructive feedback linked to the objectives.',
      'Highlight correct language, methods or values related to "' + topic + '" and unit "' + unit + '".'
    ],
    rw: [
      'Sobanura ibitekerezo by\'ingenzi bya "' + topic + '" ukoresheje ururimi rworoshye n\'ibikoresho bya REB.',
      'Erekana intambwe z\'ingenzi ku kibaho kugira ngo abanyeshuri babone uko "' + topic + '" bikora.',
      'Tegura amatsinda magufi akora umurimo cyangwa ikiganiro ku "' + topic + '".',
      'Nyura mu class ushirmikiye amatsinda, ubaze ibibazo kandi ukosore amakosa vuba.',
      'Emera amatsinda atange ikinamico, uhagure neza uhuje n\'intego.',
      'Shyira ahaboneka indimi, uburyo cyangwa indangagaciro zijyanye na "' + topic + '".'
    ],
    fr: [
      'Expliquer les notions principales de « ' + topic + ' » avec un langage simple et les supports REB.',
      'Démontrer les étapes clés au tableau pour montrer « ' + topic + ' » en pratique.',
      'Organiser des binômes ou groupes pour une tâche guidée sur « ' + topic + ' ».',
      'Circuler pour soutenir, questionner et corriger les erreurs rapidement.',
      'Faire présenter les groupes et donner un feedback précis lié aux objectifs.',
      'Souligner le langage, les méthodes ou les valeurs liées à « ' + topic + ' ».'
    ],
    sw: [
      'Eleza dhana kuu za "' + topic + '" kwa lugha rahisi na vifaa vya REB.',
      'Onyesha hatua muhimu kwenye ubao ili wanafunzi waone "' + topic + '" vitendo.',
      'Panga vikundi vidogo kwa kazi au majadiliano kuhusu "' + topic + '".',
      'Zunguka darasani kusaidia, kuuliza maswali na kusahihisha mapema.',
      'Waalike wanafunzi kuwasilisha na toa maoni yanayohusiana na malengo.',
      'Sisitiza lugha, mbinu au maadili yanayohusiana na "' + topic + '".'
    ]
  })[lang];

  var devL = ({
    en: [
      'Listen attentively and follow the explanation of "' + topic + '", taking short notes where useful.',
      'Observe the demonstration and ask questions when a step is not clear.',
      'Work actively in pairs or groups on the assigned task about "' + topic + '", sharing roles fairly.',
      'Present group ideas briefly and listen respectfully to other groups.',
      'Use feedback to improve understanding and correct their own work.',
      'Practise the key skill or idea of the lesson with support from peers and the teacher.'
    ],
    rw: [
      'Tega amatwi basobanukirwa "' + topic + '", bandike inyandiko ngufi.',
      'Reba ikinamico babaze iyo intambwe idasobanutse.',
      'Kora mu matsinda ku murimo wa "' + topic + '", bagabane inshingano.',
      'Tanga ikinamico ngufi, bashinze amatwi ku yandi matsinda.',
      'Koresha inyunganizi bagabanye ubumenyi.',
      'Iyemeze ubuhanga ngombwa bw\'isomo bafashijwe n\'inshuti n\'umwarimu.'
    ],
    fr: [
      'Écouter et suivre l\'explication de « ' + topic + ' », prendre de brèves notes.',
      'Observer la démonstration et poser des questions si besoin.',
      'Travailler en groupe sur la tâche liée à « ' + topic + ' » en partageant les rôles.',
      'Présenter brièvement et écouter les autres groupes avec respect.',
      'Utiliser le feedback pour améliorer le travail.',
      'S\'exercer sur la compétence clé de la leçon avec l\'aide des pairs et de l\'enseignant.'
    ],
    sw: [
      'Sikiliza maelezo ya "' + topic + '" na andika muhtasari.',
      'Angalia onyesho na uliza maswali inapohitajika.',
      'Fanya kazi kwa vikundi kuhusu "' + topic + '" kwa kugawana majukumu.',
      'Wasilisha kwa ufupi na wasikilize wengine kwa heshima.',
      'Tumia maoni kuboresha uelewa na kazi.',
      'Fanya mazoezi ya stadi kuu ya somo kwa msaada wa rika na mwalimu.'
    ]
  })[lang];

  var devC = ({
    en: [
      'Cooperation: sharing tasks and supporting peers while exploring "' + topic + '".',
      'Critical thinking: analysing examples and applying ideas about "' + topic + '" to new situations.',
      'Communication: explaining findings clearly to the class.',
      'Lifelong learning: practising independently and learning from feedback.'
    ],
    rw: [
      'Gukorera hamwe: kugabana imirimo no gufashanya ku "' + topic + '".',
      'Gtekereza cyimbitse: gusesengura n\'gukoresha ibitekerezo ku "' + topic + '".',
      'Itumanaho: gusobanura neza ibyavuye mu matsinda.',
      'Kwiga ubuzima bwose: kwitoza no kwiga ku nyunganizi.'
    ],
    fr: [
      'Coopération : partage des tâches sur « ' + topic + ' ».',
      'Pensée critique : analyser et appliquer les idées de « ' + topic + ' ».',
      'Communication : présenter clairement les résultats.',
      'Apprentissage tout au long de la vie : s\'exercer et apprendre du feedback.'
    ],
    sw: [
      'Ushirikiano: kugawana kazi kuhusu "' + topic + '".',
      'Fikra kritikal: kuchambua na kutumia mawazo ya "' + topic + '".',
      'Mawasiliano: kueleza matokeo kwa uwazi.',
      'Kujifunza maisha yote: kujizoeza na kujifunza kutoka maoni.'
    ]
  })[lang];

  var concT = ({
    en: [
      'Guide learners to summarise the key points of "' + topic + '" in their own words.',
      'Check understanding with a short oral or written task linked to the objectives.',
      'Give clear, manageable homework or practice related to unit "' + unit + '".',
      'Preview how the next lesson will build on today\'s learning about "' + topic + '".'
    ],
    rw: [
      'Bayobore bagusobanura mu magambo yabo ingingo z\'ingenzi za "' + topic + '".',
      'Suzuma ubumenyi hakoreshejwe ikibazo cy\'igufi gihuye n\'intego.',
      'Tanga umurimo wo mu rugo woroshye uhuye n\'igice "' + unit + '".',
      'Vuga briefly isomo rikurikira rizubaka ku "' + topic + '".'
    ],
    fr: [
      'Faire résumer par les élèves les points clés de « ' + topic + ' ».',
      'Vérifier la compréhension par une brève tâche orale ou écrite.',
      'Donner un devoir simple lié à l\'unité « ' + unit + ' ».',
      'Annoncer le lien avec la prochaine leçon sur « ' + topic + ' ».'
    ],
    sw: [
      'Waongoze wanafunzi kufupisha mambo makuu ya "' + topic + '".',
      'Kagua uelewa kwa kazi fupi ya mdomo au maandishi.',
      'Toa kazi ya nyumbani rahisi inayohusiana na kitengo "' + unit + '".',
      'Tangaza jinsi somo lijalo litakavyojenga juu ya "' + topic + '".'
    ]
  })[lang];

  var concL = ({
    en: [
      'Contribute to the class summary using correct ideas about "' + topic + '".',
      'Complete the short check activity carefully and honestly.',
      'Record homework instructions and ask final questions if needed.',
      'Reflect briefly on what they can now do better related to the lesson.'
    ],
    rw: [
      'Bagire uruhare mu gusouza isomo bakoresheje ibitekerezo by\'ukuri ku "' + topic + '".',
      'Kora ikibazo cyo gusuzuma neza kandi mu kuri.',
      'Andika umurimo wo mu rugo, babaze ibibazo bya nyuma.',
      'Tekereze ku cyo ubu bashobora gukora neza kurushaho.'
    ],
    fr: [
      'Participer au résumé avec des idées justes sur « ' + topic + ' ».',
      'Réaliser la tâche de vérification avec soin.',
      'Noter le devoir et poser les dernières questions.',
      'Réfléchir brièvement à ce qu\'ils savent mieux faire.'
    ],
    sw: [
      'Shiriki katika muhtasari kwa mawazo sahihi kuhusu "' + topic + '".',
      'Kamilisha kazi fupi ya ukaguzi kwa makini.',
      'Andika maelekezo ya kazi ya nyumbani na uliza maswali ya mwisho.',
      'Tafakari kwa ufupi wanachoweza kufanya vizuri sasa.'
    ]
  })[lang];

  var concC = ({
    en: [
      'Communication: summarising learning about "' + topic + '" in own words.',
      'Lifelong learning: taking responsibility for practice at home.',
      'Inclusive education: ensuring every learner leaves with a clear next step.'
    ],
    rw: [
      'Itumanaho: gusouza isomo mu magambo yabo ku "' + topic + '".',
      'Kwiga ubuzima bwose: kwiyalamira imyitozo yo mu rugo.',
      'Uburezi bw\'abantu bose: umunyeshuri wese asiga afite intambwe ikurikira.'
    ],
    fr: [
      'Communication : résumer l\'apprentissage sur « ' + topic + ' ».',
      'Apprentissage tout au long de la vie : responsabilité du travail à la maison.',
      'Éducation inclusive : chaque apprenant repart avec une prochaine étape claire.'
    ],
    sw: [
      'Mawasiliano: kufupisha kujifunza kuhusu "' + topic + '".',
      'Kujifunza maisha yote: kujwajibika kwa mazoezi nyumbani.',
      'Elimu jumuishi: kila mwanafunzi anaondoka na hatua ijayo wazi.'
    ]
  })[lang];

  var selfDraft = ({
    en: 'Most learners achieved the objectives on "' + topic + '". Group work on unit "' + unit + '" was active and inclusive. Next lesson I will support slower learners with extra examples and give stronger groups a short extension task so the class remains fully engaged.',
    rw: 'Abanyeshuri benshi bageze ku ntego z\'isomo "' + topic + '". Akazi k\'amatsinda ku gice "' + unit + '" kari gakora. Mu isomo rikurikira nzashyigikira abanyeshuri basuye n\'ibindi byokureba kandi nkeza amatsinda ashoboye umurimo w\'inyongera.',
    fr: 'La plupart des élèves ont atteint les objectifs sur « ' + topic + ' ». Le travail de groupe sur l\'unité « ' + unit + ' » était actif. La prochaine fois, je soutiendrai davantage les élèves en difficulté et proposerai une tâche d\'extension aux groupes plus avancés.',
    sw: 'Wanafunzi wengi walifika malengo kuhusu "' + topic + '". Kazi za vikundi kwenye kitengo "' + unit + '" zilikuwa hai. Somo lijalo nitasaidia wanaohitaji mifano zaidi na kuwapa vikundi imara kazi ya ziada.'
  })[lang];

  return {
    key: keyCompetence,
    obj: objective,
    mats: mats,
    refs: references,
    activityNote: activityNote,
    introT: introT,
    introL: introL,
    introC: introC,
    devT: devT,
    devL: devL,
    devC: devC,
    concT: concT,
    concL: concL,
    concC: concC,
    selfDraft: selfDraft
  };
}

function onLanguageChange() {
  var lang = document.getElementById('languageSelect').value || 'en';
  // If a plan is already generated, rebuild full content in the new language
  if (LAST) {
    LAST.lang = lang;
    // Force re-auto content in new language (clear optional fields that were auto-filled)
    var data = Object.assign({}, LAST);
    data.lang = lang;
    // Re-run smart content in new language
    // Prefer key competence already written in the form (REB text) — do not overwrite
    var formKC = '';
    try { formKC = (document.getElementById('keyCompetence') && document.getElementById('keyCompetence').value.trim()) || ''; } catch (e) {}
    var smart = buildSmartContent({
      lessonTitle: data.lessonTitle,
      unitTitle: data.unitTitle,
      subject: data.subject,
      class: data.class,
      lang: lang,
      keyCompetence: formKC,
      objective: '',
      materials: '',
      references: ''
    });
    // Keep REB/form key competence as written; only auto-fill if empty
    data.keyCompetence = formKC.length >= 20 ? formKC : smart.key;
    data.objective = smart.obj;
    data.materials = smart.mats;
    data.references = smart.refs;
    data.activityNote = smart.activityNote;
    data.introT = smart.introT; data.introL = smart.introL; data.introC = smart.introC;
    data.devT = smart.devT; data.devL = smart.devL; data.devC = smart.devC;
    data.concT = smart.concT; data.concL = smart.concL; data.concC = smart.concC;
    // Keep self-eval if user wrote custom; else new draft
    if (!data.selfEvaluation || data.selfEvaluation === (LAST._lastSelfDraft || '')) {
      data.selfEvaluation = smart.selfDraft;
    }
    data._lastSelfDraft = smart.selfDraft;
    LAST = data;
    try {
      var se = document.getElementById('selfEvalInput');
      if (se) se.value = data.selfEvaluation || '';
    } catch (e) {}
    reRenderFromLAST();
    try { localStorage.setItem('cbc_lp_last', JSON.stringify(LAST)); } catch (e2) {}
  }
}


function onSpecialNeedsChange() {
  var sel = document.getElementById('specialNeeds');
  var det = document.getElementById('specialNeedsDetail');
  if (!sel || !det) return;
  if (sel.value === 'None') det.placeholder = 'No special needs — optional note';
  else if (sel.value === 'Other (specify below)') det.placeholder = 'Describe the special need and number of learners';
  else det.placeholder = 'Number of learners: e.g. 2 learners with ' + sel.value;
}
function getSpecialNeedsValue() {
  var sel = (document.getElementById('specialNeeds') || {}).value || 'None';
  var det = ((document.getElementById('specialNeedsDetail') || {}).value || '').trim();
  if (sel === 'None' && !det) return 'None';
  if (sel === 'None') return det;
  if (sel === 'Other (specify below)') return det || 'Other special educational needs';
  return det ? (sel + ' — ' + det) : sel;
}
function getLocationValue() {
  return ((document.getElementById('location') || {}).value || 'Classroom').trim();
}

function generateLessonPlan() {
  var t = T[document.getElementById('languageSelect').value] || T.en;
  var classBase = val('class');
  var classGroup = val('classGroup');
  var classCombined = classBase + (classGroup ? ' ' + classGroup : '');
  var data = {
    schoolName: val('schoolName'), teacherName: val('teacherName'), term: val('term'),
    date: val('date'), subject: getSubjectValue(), class: classCombined, level: val('level'),
    classBase: classBase, classGroup: classGroup,
    unitNo: val('unitNo'), lessonNo: val('lessonNo'), totalLessons: val('totalLessons'),
    duration: val('duration'), classSize: val('classSize'),
    unitTitle: val('unitTitle'), lessonTitle: val('lessonTitle'),
    keyCompetence: val('keyCompetence'), objective: val('objective'),
    specialNeeds: (typeof getSpecialNeedsValue === 'function' ? getSpecialNeedsValue() : val('specialNeeds')) || t.none,
    location: (typeof getLocationValue === 'function' ? getLocationValue() : val('location')) || t.classroom,
    materials: val('materials'), references: val('references'),
    lang: document.getElementById('languageSelect').value || 'en'
  };

  // Sensible defaults so generation is not blocked by optional numbers
  if (!data.duration) {
    data.duration = '40';
    try { document.getElementById('duration').value = '40'; } catch (e) {}
  }
  if (!data.classSize) {
    data.classSize = '40';
    try { document.getElementById('classSize').value = '40'; } catch (e) {}
  }
  if (!data.unitNo && data.unitTitle) {
    data.unitNo = '1';
    try { document.getElementById('unitNo').value = '1'; } catch (e) {}
  }
  if (!data.lessonNo && data.lessonTitle) {
    data.lessonNo = '1';
    try { document.getElementById('lessonNo').value = '1'; } catch (e) {}
  }
  if (!data.totalLessons) {
    data.totalLessons = '5';
    try { document.getElementById('totalLessons').value = '5'; } catch (e) {}
  }

  var missing = [];
  if (!data.schoolName) missing.push('School Name');
  if (!data.teacherName) missing.push("Teacher's Name");
  if (!data.term) missing.push('Term');
  if (!data.date) missing.push('Date');
  if (!data.level) missing.push('Level');
  if (!classBase) missing.push('Class');
  if (!classGroup) missing.push('Class group');
  if (!data.subject) missing.push('Subject');
  if (!data.unitTitle) missing.push('Unit title');
  if (!data.lessonTitle) missing.push('Lesson title');
  if (!data.unitNo) missing.push('Unit No');
  if (!data.lessonNo) missing.push('Lesson No');
  if (!data.totalLessons) missing.push('Total lessons');
  if (!data.duration) missing.push('Duration');
  if (!data.classSize) missing.push('Class size');
  if (missing.length) {
    var em = document.getElementById('errorMessage');
    if (em) em.textContent = 'Please fill: ' + missing.join(', ');
    alert('Please fill required fields: ' + missing.join(', '));
    return;
  }
  var em2 = document.getElementById('errorMessage');
  if (em2) em2.textContent = '';

  try { // GENERATE_SAFE
  var total = Number(data.duration) || 40;
  var introM = Math.min(8, Math.max(6, Math.round(total * 0.18)));
  var concM = Math.min(9, Math.max(6, Math.round(total * 0.18)));
  var devM = Math.max(12, total - introM - concM);

  var smart = buildSmartContent(data);
  // Keep values written in the form fields (REB key competence, custom objectives, etc.)
  // Only auto-fill when the field is empty or very short
  var formKC = (data.keyCompetence || '').trim();
  var formObj = (data.objective || '').trim();
  var formMats = (data.materials || '').trim();
  var formRefs = (data.references || '').trim();

  data.keyCompetence = formKC.length >= 20 ? formKC : smart.key;
  data.objective = formObj.length >= 25 ? formObj : smart.obj;
  data.materials = formMats.length >= 15 ? formMats : smart.mats;
  data.references = formRefs.length >= 15 ? formRefs : smart.refs;

  // Sync back to form so user sees what will appear on the plan
  try {
    document.getElementById('keyCompetence').value = data.keyCompetence;
    if (!formObj) document.getElementById('objective').value = data.objective;
    if (!formMats) document.getElementById('materials').value = data.materials;
    if (!formRefs) document.getElementById('references').value = data.references;
  } catch (eSync) {}

  if (!val('specialNeeds')) data.specialNeeds = t.none;
  if (!val('location')) data.location = t.classroom;
  data.introM = introM; data.devM = devM; data.concM = concM;
  data.introT = smart.introT; data.introL = smart.introL; data.introC = smart.introC;
  data.devT = smart.devT; data.devL = smart.devL; data.devC = smart.devC;
  data.concT = smart.concT; data.concL = smart.concL; data.concC = smart.concC;
  data.activityNote = smart.activityNote || t.activityNote;
  data.selfEvaluation = (document.getElementById('selfEvalInput') && document.getElementById('selfEvalInput').value.trim())
    ? document.getElementById('selfEvalInput').value.trim()
    : (smart.selfDraft || '');
  data.lang = document.getElementById('languageSelect').value || 'en';
  data._lastSelfDraft = smart.selfDraft || '';
  LAST = data;

  var lessonNoText = data.lessonNo + ' ' + (t.lessonOf || 'of') + ' ' + data.totalLessons;
  var objHtml = esc(data.objective).replace(/\n/g, '<br>');

  var html = ''
    + '<div class="title">' + t.lessonPlan + '</div>'
    + '<div class="meta"><b>' + t.schoolName + ':</b> ' + esc(data.schoolName)
    + ' &nbsp;&nbsp;&nbsp; <b>' + t.teacherName + ':</b> ' + esc(data.teacherName) + '</div>'
    + '<table>'
    + '<tr class="hdr"><td>' + t.term + '</td><td>' + t.date + '</td><td>' + t.subject + '</td><td>' + t.class + '</td>'
    + '<td>' + t.unitNo + '</td><td>' + t.lessonNo + '</td><td>' + t.duration + '</td><td>' + t.classSize + '</td></tr>'
    + '<tr><td>' + esc(data.term) + '</td><td>' + esc(data.date) + '</td><td>' + esc(data.subject) + '</td><td>' + esc(data.class) + '</td>'
    + '<td>' + esc(data.unitNo) + '</td><td>' + esc(lessonNoText) + '</td><td>' + esc(data.duration) + ' min</td><td>' + esc(data.classSize) + '</td></tr>'
    + '</table>'
    + '<table>'
    + '<tr><td class="lab">' + t.specialNeeds + '</td><td>' + esc(data.specialNeeds) + '</td></tr>'
    + '<tr><td class="lab">' + t.unitTitle + '</td><td>' + esc(data.unitTitle) + '</td></tr>'
    + '<tr><td class="lab">' + t.keyCompetence + '</td><td>' + esc(data.keyCompetence) + '</td></tr>'
    + '<tr><td class="lab">' + t.lessonTitle + '</td><td>' + esc(data.lessonTitle) + '</td></tr>'
    + '<tr><td class="lab">' + t.objective + '</td><td>' + objHtml + '</td></tr>'
    + '<tr><td class="lab">' + t.planClass + '</td><td>' + esc(data.location) + '</td></tr>'
    + '<tr><td class="lab">' + t.materials + '</td><td>' + esc(data.materials) + '</td></tr>'
    + '<tr><td class="lab">' + t.references + '</td><td>' + esc(data.references).replace(/\n/g,'<br>') + '</td></tr>'
    + '</table>'
    + '<table>'
    + '<tr class="hdr">'
    + '<td style="width:15%;text-align:center;vertical-align:middle" rowspan="2">' + t.timing + '</td>'
    + '<td colspan="2" style="text-align:center">' + t.activities
    + '<div style="font-weight:400;font-style:italic;margin-top:3px;font-size:0.92em">' + esc(data.activityNote || t.activityNote) + '</div></td>'
    + '<td style="width:28%;text-align:center;vertical-align:middle" rowspan="2">' + t.competences + '</td></tr>'
    + '<tr class="hdr"><td style="width:28.5%;text-align:center">' + t.teacherAct + '</td><td style="width:28.5%;text-align:center">' + t.learnerAct + '</td></tr>'
    + '<tr><td class="lab">' + t.introduction + ' (' + introM + ' ' + (t.minutes || 'min') + ')</td>'
    + '<td>' + bullets(data.introT) + '</td><td>' + bullets(data.introL) + '</td><td>' + bullets(data.introC) + '</td></tr>'
    + '<tr><td class="lab">' + t.development + ' (' + devM + ' ' + (t.minutes || 'min') + ')</td>'
    + '<td>' + bullets(data.devT) + '</td><td>' + bullets(data.devL) + '</td><td>' + bullets(data.devC) + '</td></tr>'
    + '<tr><td class="lab">' + t.conclusion + ' (' + concM + ' ' + (t.minutes || 'min') + ')'
    + (t.conclusionNote ? '<br><span style="font-weight:600;font-size:0.9em">' + t.conclusionNote + '</span>' : '')
    + '</td>'
    + '<td>' + bullets(data.concT) + '</td><td>' + bullets(data.concL) + '</td><td>' + bullets(data.concC) + '</td></tr>'
    + '<tr><td class="lab">' + t.selfEval + '</td><td colspan="3" id="selfEvalCell">' + (data.selfEvaluation ? esc(data.selfEvaluation).replace(/\n/g,'<br>') : '&nbsp;<br>&nbsp;<br>&nbsp;') + '</td></tr>'
    + '</table>'
    + '<div class="plan-footer"><span style="float:left;font-weight:500;color:#94a3b8;font-size:8pt">' + esc(data.date || '') + '</span>SL Lesson plan</div>';

  document.getElementById('lessonPlanContent').innerHTML = html;
  var rc = document.getElementById('resultContainer');
  if (rc) {
    rc.classList.add('show');
    rc.style.display = 'block';
  }
  var se = document.getElementById('selfEvalInput');
  if (se) se.value = data.selfEvaluation || '';
  try {
    var ss = document.getElementById('selfEvalSelect');
    if (ss && ss.value) onSelfEvalSelect();
  } catch (e1) {}
  document.getElementById('resultContainer').scrollIntoView({ behavior: 'smooth' });
    try { localStorage.setItem('cbc_lp_last', JSON.stringify(data)); } catch (e) {}
  } catch (genErr) {
    console.error(genErr);
    var em3 = document.getElementById('errorMessage');
    if (em3) em3.textContent = 'Generate error: ' + (genErr && genErr.message ? genErr.message : genErr);
    alert('Could not generate lesson plan: ' + (genErr && genErr.message ? genErr.message : genErr));
  }
}function applySelfEvaluation() {
  var txt = (document.getElementById('selfEvalInput').value || '').trim();
  if (!LAST) { alert('Generate a lesson plan first.'); return; }
  LAST.selfEvaluation = txt;
  var cell = document.getElementById('selfEvalCell');
  if (cell) {
    cell.innerHTML = txt ? esc(txt).replace(/\n/g, '<br>') : '&nbsp;<br>&nbsp;<br>&nbsp;';
  }
  try { localStorage.setItem('cbc_lp_last', JSON.stringify(LAST)); } catch (e) {}
  alert(txt ? 'Self-evaluation applied to the lesson plan. Downloads will include it.' : 'Self-evaluation cleared.');
}


function val(id) { return (document.getElementById(id).value || '').trim(); }

function downloadSelected() {
  if (!LAST) { alert('Generate a lesson plan first.'); return; }
  if (document.getElementById('dlPdf').checked) exportPDF();
  if (document.getElementById('dlWord').checked) exportWord();
  if (document.getElementById('dlExcel').checked) exportExcel();
  if (document.getElementById('dlCsv').checked) exportCSV();
}


var EDIT_MODE = false;

function linesToArr(txt) {
  return String(txt || '').split(/\n+/).map(function (s) { return s.replace(/^[-•\s]+/, '').trim(); }).filter(Boolean);
}
function arrToLines(arr) {
  return (arr || []).join('\n');
}

function toggleEditMode() {
  var el = document.getElementById('lessonPlanContent');
  if (!el || !LAST) { alert('Generate a lesson plan first.'); return; }
  EDIT_MODE = !EDIT_MODE;
  el.contentEditable = EDIT_MODE ? 'true' : 'false';
  el.classList.toggle('editing', EDIT_MODE);
  if (!EDIT_MODE) {
    // leaving edit mode: keep HTML as-is for Word/print; PDF uses LAST so offer sync tip
    alert('Edit mode off. For PDF/Excel, use "Structured fields" → Apply, or keep Word/Print from the edited view.');
  } else {
    el.focus();
  }
}

function openStructuredEdit() {
  if (!LAST) { alert('Generate a lesson plan first.'); return; }
  var box = document.getElementById('structuredEdit');
  box.style.display = box.style.display === 'none' ? 'block' : 'none';
  if (box.style.display === 'none') return;
  document.getElementById('editKeyCompetence').value = LAST.keyCompetence || '';
  document.getElementById('editObjective').value = LAST.objective || '';
  document.getElementById('editMaterials').value = LAST.materials || '';
  document.getElementById('editReferences').value = LAST.references || '';
  document.getElementById('editIntroT').value = arrToLines(LAST.introT);
  document.getElementById('editIntroL').value = arrToLines(LAST.introL);
  document.getElementById('editIntroC').value = arrToLines(LAST.introC);
  document.getElementById('editDevT').value = arrToLines(LAST.devT);
  document.getElementById('editDevL').value = arrToLines(LAST.devL);
  document.getElementById('editDevC').value = arrToLines(LAST.devC);
  document.getElementById('editConcT').value = arrToLines(LAST.concT);
  document.getElementById('editConcL').value = arrToLines(LAST.concL);
  document.getElementById('editConcC').value = arrToLines(LAST.concC);
  document.getElementById('editSelfEval').value = LAST.selfEvaluation || '';
}

function applyStructuredEdit() {
  if (!LAST) return;
  LAST.keyCompetence = document.getElementById('editKeyCompetence').value.trim();
  LAST.objective = document.getElementById('editObjective').value.trim();
  LAST.materials = document.getElementById('editMaterials').value.trim();
  LAST.references = document.getElementById('editReferences').value.trim();
  LAST.introT = linesToArr(document.getElementById('editIntroT').value);
  LAST.introL = linesToArr(document.getElementById('editIntroL').value);
  LAST.introC = linesToArr(document.getElementById('editIntroC').value);
  LAST.devT = linesToArr(document.getElementById('editDevT').value);
  LAST.devL = linesToArr(document.getElementById('editDevL').value);
  LAST.devC = linesToArr(document.getElementById('editDevC').value);
  LAST.concT = linesToArr(document.getElementById('editConcT').value);
  LAST.concL = linesToArr(document.getElementById('editConcL').value);
  LAST.concC = linesToArr(document.getElementById('editConcC').value);
  LAST.selfEvaluation = document.getElementById('editSelfEval').value.trim();
  try {
    var se = document.getElementById('selfEvalInput');
    if (se) se.value = LAST.selfEvaluation;
  } catch (e) {}
  reRenderFromLAST();
  try { localStorage.setItem('cbc_lp_last', JSON.stringify(LAST)); } catch (e2) {}
  alert('Edits applied. Downloads will use the updated plan.');
}

function reRenderFromLAST() {
  if (!LAST) return;
  // Rebuild HTML using same path as generate — call internal renderer if available
  try {
    if (typeof renderLessonPlanHTML === 'function') {
      renderLessonPlanHTML(LAST);
      return;
    }
  } catch (e) {}
  // Fallback: re-run generate with form values already filled — store and display from LAST manually
  var t = T[LAST.lang || 'en'] || T.en;
  var data = LAST;
  var introM = data.introM, devM = data.devM, concM = data.concM;
  function bullets(arr) {
    if (!arr || !arr.length) return '&nbsp;';
    return '<ul>' + arr.map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') + '</ul>';
  }
  function esc(s) {
    return String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }
  var lessonNoText = data.lessonNo + ' ' + (t.lessonOf || 'of') + ' ' + data.totalLessons;
  var objHtml = esc(data.objective).replace(/\n/g, '<br>');
  var html = ''
    + '<div class="title">' + t.lessonPlan + '</div>'
    + '<div class="meta"><b>' + t.schoolName + ':</b> ' + esc(data.schoolName)
    + ' &nbsp;&nbsp;&nbsp; <b>' + t.teacherName + ':</b> ' + esc(data.teacherName) + '</div>'
    + '<table>'
    + '<tr class="hdr"><td>' + t.term + '</td><td>' + t.date + '</td><td>' + t.subject + '</td><td>' + t.class + '</td>'
    + '<td>' + t.unitNo + '</td><td>' + t.lessonNo + '</td><td>' + t.duration + '</td><td>' + t.classSize + '</td></tr>'
    + '<tr><td>' + esc(data.term) + '</td><td>' + esc(data.date) + '</td><td>' + esc(data.subject) + '</td><td>' + esc(data.class) + '</td>'
    + '<td>' + esc(data.unitNo) + '</td><td>' + esc(lessonNoText) + '</td><td>' + esc(data.duration) + ' min</td><td>' + esc(data.classSize) + '</td></tr>'
    + '</table>'
    + '<table>'
    + '<tr><td class="lab">' + t.specialNeeds + '</td><td>' + esc(data.specialNeeds) + '</td></tr>'
    + '<tr><td class="lab">' + t.unitTitle + '</td><td>' + esc(data.unitTitle) + '</td></tr>'
    + '<tr><td class="lab">' + t.keyCompetence + '</td><td>' + esc(data.keyCompetence) + '</td></tr>'
    + '<tr><td class="lab">' + t.lessonTitle + '</td><td>' + esc(data.lessonTitle) + '</td></tr>'
    + '<tr><td class="lab">' + t.objective + '</td><td>' + objHtml + '</td></tr>'
    + '<tr><td class="lab">' + t.planClass + '</td><td>' + esc(data.location) + '</td></tr>'
    + '<tr><td class="lab">' + t.materials + '</td><td>' + esc(data.materials) + '</td></tr>'
    + '<tr><td class="lab">' + t.references + '</td><td>' + esc(data.references).replace(/\n/g,'<br>') + '</td></tr>'
    + '</table>'
    + '<table>'
    + '<tr class="hdr">'
    + '<td style="width:15%;text-align:center;vertical-align:middle" rowspan="2">' + t.timing + '</td>'
    + '<td colspan="2" style="text-align:center">' + t.activities
    + '<div style="font-weight:400;font-style:italic;margin-top:3px;font-size:0.92em">' + esc(data.activityNote || t.activityNote) + '</div></td>'
    + '<td style="width:28%;text-align:center;vertical-align:middle" rowspan="2">' + t.competences + '</td></tr>'
    + '<tr class="hdr"><td style="width:28.5%;text-align:center">' + t.teacherAct + '</td><td style="width:28.5%;text-align:center">' + t.learnerAct + '</td></tr>'
    + '<tr><td class="lab">' + t.introduction + ' (' + introM + ' ' + (t.minutes || 'min') + ')</td>'
    + '<td>' + bullets(data.introT) + '</td><td>' + bullets(data.introL) + '</td><td>' + bullets(data.introC) + '</td></tr>'
    + '<tr><td class="lab">' + t.development + ' (' + devM + ' ' + (t.minutes || 'min') + ')</td>'
    + '<td>' + bullets(data.devT) + '</td><td>' + bullets(data.devL) + '</td><td>' + bullets(data.devC) + '</td></tr>'
    + '<tr><td class="lab">' + t.conclusion + ' (' + concM + ' ' + (t.minutes || 'min') + ')'
    + (t.conclusionNote ? '<br><span style="font-weight:600;font-size:0.9em">' + t.conclusionNote + '</span>' : '')
    + '</td>'
    + '<td>' + bullets(data.concT) + '</td><td>' + bullets(data.concL) + '</td><td>' + bullets(data.concC) + '</td></tr>'
    + '<tr><td class="lab">' + t.selfEval + '</td><td colspan="3" id="selfEvalCell">' + (data.selfEvaluation ? esc(data.selfEvaluation).replace(/\n/g,'<br>') : '&nbsp;<br>&nbsp;<br>&nbsp;') + '</td></tr>'
    + '</table>'
    + '<div class="plan-footer"><span style="float:left;font-weight:500;color:#94a3b8;font-size:8pt">' + esc(data.date || '') + '</span>SL Lesson plan</div>';
  document.getElementById('lessonPlanContent').innerHTML = html;
  EDIT_MODE = false;
  document.getElementById('lessonPlanContent').contentEditable = 'false';
  document.getElementById('lessonPlanContent').classList.remove('editing');
}


function exportPDF() {
  if (!LAST) { alert('Generate a lesson plan first.'); return; }
  var d = LAST;
  try {
    var se = document.getElementById('selfEvalInput');
    if (se) d.selfEvaluation = (se.value || '').trim();
  } catch (e) {}

  if (typeof window.jspdf === 'undefined' || !window.jspdf.jsPDF) {
    alert('PDF library not loaded. Check your internet connection and refresh.');
    return;
  }

  var jsPDF = window.jspdf.jsPDF;
  var doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4', compress: true });
  var pageW = doc.internal.pageSize.getWidth();
  var pageH = doc.internal.pageSize.getHeight();
  var margin = 5.5;
  var footerH = 6;

  function take(arr, n) {
    arr = arr || [];
    if (arr.length <= n) return arr;
    return arr.slice(0, n);
  }

  function join(a) {
    return (a || []).map(function (x) { return '- ' + x; }).join('\n');
  }

  // Content tuned for ONE A4 page: full meaning, slightly capped lists
  // Match sample density so PDF always fits ONE A4 page while keeping full structure
  var introT = take(d.introT, 5);
  var introL = take(d.introL, 5);
  var introC = take(d.introC, 3);
  var devT = take(d.devT, 6);
  var devL = take(d.devL, 6);
  var devC = take(d.devC, 4);
  var concT = take(d.concT, 4);
  var concL = take(d.concL, 4);
  var concC = take(d.concC, 3);
  var selfTxt = d.selfEvaluation || '';

  function clearDoc() {
    while (doc.internal.getNumberOfPages() > 1) {
      doc.deletePage(doc.internal.getNumberOfPages());
    }
    doc.setPage(1);
    doc.setFillColor(255, 255, 255);
    doc.rect(0, 0, pageW, pageH, 'F');
  }

  function drawAtScale(scale) {
    clearDoc();
    var fs = Math.max(8.4, Math.min(12.5, 10.8 * scale));
    var fsTitle = Math.max(13, Math.min(17, 15.2 * scale));
    var fsMeta = Math.max(10, Math.min(13, 11.5 * scale));
    var pad = Math.max(0.55, 1.0 * scale);
    var y = margin + 1.5;

    var Lt = T[d.lang || 'en'] || T.en;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(fsTitle);
    doc.setTextColor(0, 0, 0);
    doc.text(String(Lt.lessonPlan || 'LESSON PLAN'), pageW / 2, y, { align: 'center' });
    y += 4.5 * scale + 0.8;

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(fsMeta);
    doc.text(String(Lt.schoolName || 'School Name') + ': ' + String(d.schoolName || ''), margin, y);
    doc.text(String(Lt.teacherName || "Teacher's Name") + ': ' + String(d.teacherName || ''), pageW / 2 + 1, y);
    y += 2.2;

    var common = {
      theme: 'grid',
      styles: {
        fontSize: fs,
        cellPadding: pad,
        lineColor: [0, 0, 0],
        lineWidth: 0.2,
        textColor: [0, 0, 0],
        overflow: 'linebreak',
        valign: 'top',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [200, 205, 210],
        textColor: [0, 0, 0],
        fontStyle: 'bold',
        fontSize: fs
      },
      margin: { left: margin, right: margin, bottom: margin + footerH },
      pageBreak: 'avoid',
      rowPageBreak: 'avoid',
      showHead: 'everyPage'
    };

    doc.autoTable(Object.assign({}, common, {
      startY: y,
      head: [[Lt.term, Lt.date, Lt.subject, Lt.class, Lt.unitNo, Lt.lessonNo, Lt.duration, Lt.classSize]],
      body: [[
        d.term, d.date, d.subject, d.class, d.unitNo,
        d.lessonNo + ' ' + (Lt.lessonOf || 'of') + ' ' + d.totalLessons,
        d.duration + ' ' + (Lt.minutes || 'min'), d.classSize
      ]],
      styles: Object.assign({}, common.styles, { fontSize: Math.max(6, fs - 0.4), cellPadding: pad * 0.85, halign: 'center' })
    }));
    y = doc.lastAutoTable.finalY;

    var infoBody = [
      [Lt.specialNeeds, d.specialNeeds],
      [Lt.unitTitle, d.unitTitle],
      [Lt.keyCompetence, d.keyCompetence],
      [Lt.lessonTitle, d.lessonTitle],
      [Lt.objective, d.objective],
      [Lt.planClass, d.location],
      [Lt.materials, d.materials],
      [Lt.references, d.references]
    ];

    doc.autoTable(Object.assign({}, common, {
      startY: y,
      body: infoBody,
      columnStyles: {
        0: { cellWidth: 48, fillColor: [232, 234, 237], fontStyle: 'bold', fontSize: Math.max(6, fs - 0.3) },
        1: { cellWidth: pageW - margin * 2 - 48 }
      },
      styles: Object.assign({}, common.styles, { fontSize: Math.max(6.2, fs - 0.2) })
    }));
    y = doc.lastAutoTable.finalY;

    var usableW = pageW - margin * 2;
    var c0 = 28;
    var cRest = (usableW - c0) / 3;
    var actNote = d.activityNote || ('Engaging students through explanation, group work, and presentations to explore ' + (d.lessonTitle || 'the topic') + '.');

    // Exact REB-style header: Timing | Description of teaching and learning activities | Generic competences...
    doc.autoTable(Object.assign({}, common, {
      startY: y,
      head: [
        [
          { content: Lt.timing, rowSpan: 3, styles: { halign: 'center', valign: 'middle', fontStyle: 'bold', fillColor: [200, 205, 210] } },
          { content: Lt.activities, colSpan: 2, styles: { halign: 'center', fontStyle: 'bold', fillColor: [200, 205, 210] } },
          { content: Lt.competences, rowSpan: 3, styles: { halign: 'center', valign: 'middle', fontStyle: 'bold', fillColor: [200, 205, 210] } }
        ],
        [
          { content: actNote, colSpan: 2, styles: { halign: 'center', fontStyle: 'italic', fontSize: Math.max(5.8, fs - 1.2), fillColor: [245, 245, 245], textColor: [40, 40, 40] } }
        ],
        [
          { content: Lt.teacherAct, styles: { halign: 'center', fontStyle: 'bold', fillColor: [200, 205, 210] } },
          { content: Lt.learnerAct, styles: { halign: 'center', fontStyle: 'bold', fillColor: [200, 205, 210] } }
        ]
      ],
      body: [
        [
          { content: Lt.introduction + '\n(' + d.introM + ' ' + (Lt.minutes||'min') + ')', styles: { fontStyle: 'bold', fillColor: [232, 234, 237] } },
          join(introT),
          join(introL),
          join(introC)
        ],
        [
          { content: Lt.development + '\n(' + d.devM + ' ' + (Lt.minutes||'min') + ')', styles: { fontStyle: 'bold', fillColor: [232, 234, 237] } },
          join(devT),
          join(devL),
          join(devC)
        ],
        [
          { content: Lt.conclusion + '\n(' + d.concM + ' ' + (Lt.minutes||'min') + ')', styles: { fontStyle: 'bold', fillColor: [232, 234, 237] } },
          join(concT),
          join(concL),
          join(concC)
        ],
        [
          { content: Lt.selfEval, styles: { fontStyle: 'bold', fillColor: [232, 234, 237] } },
          { content: selfTxt || ' ', colSpan: 3 }
        ]
      ],
      columnStyles: {
        0: { cellWidth: c0 },
        1: { cellWidth: cRest },
        2: { cellWidth: cRest },
        3: { cellWidth: cRest }
      },
      styles: Object.assign({}, common.styles, {
        fontSize: Math.max(6, fs - 0.5),
        cellPadding: pad * 0.75
      })
    }));

    return doc.lastAutoTable.finalY;

  }

  // Fit completely on ONE A4 page — table must touch the footer line
  var targetBottom = pageH - margin - footerH; // content should end at/near this Y
  var lo = 0.58, hi = 1.18, best = 0.88, bestY = 0;
  for (var i = 0; i < 16; i++) {
    var mid = (lo + hi) / 2;
    var yEnd = drawAtScale(mid);
    var pages = doc.internal.getNumberOfPages();
    var fits = (pages <= 1) && (yEnd <= targetBottom + 0.4);
    if (fits) {
      best = mid;
      bestY = yEnd;
      lo = mid;
    } else {
      hi = mid;
    }
  }
  bestY = drawAtScale(best);
  while (doc.internal.getNumberOfPages() > 1) {
    doc.deletePage(doc.internal.getNumberOfPages());
  }
  if (bestY > targetBottom + 0.8) {
    best *= 0.94;
    bestY = drawAtScale(best);
    while (doc.internal.getNumberOfPages() > 1) {
      doc.deletePage(doc.internal.getNumberOfPages());
    }
  }
  // Aggressively fill remaining blank space so the table touches the footer line
  var empty = targetBottom - bestY;
  if (empty > 2.5 && best < 1.28) {
    var tryScale = Math.min(1.22, best * (1 + Math.min(0.18, empty / 140)));
    var y2 = drawAtScale(tryScale);
    if (doc.internal.getNumberOfPages() <= 1 && y2 <= targetBottom + 0.5) {
      best = tryScale;
      bestY = y2;
    } else {
      // binary fine-tune upward until just under footer
      var lo2 = best, hi2 = tryScale;
      for (var j = 0; j < 10; j++) {
        var mid2 = (lo2 + hi2) / 2;
        var y3 = drawAtScale(mid2);
        if (doc.internal.getNumberOfPages() <= 1 && y3 <= targetBottom + 0.4) {
          lo2 = mid2;
          best = mid2;
          bestY = y3;
        } else {
          hi2 = mid2;
        }
      }
      bestY = drawAtScale(best);
      while (doc.internal.getNumberOfPages() > 1) doc.deletePage(doc.internal.getNumberOfPages());
    }
  }

  // Footer line drawn immediately under the table (table touches the line)
  doc.setPage(1);
  var footerLineY = Math.min(pageH - margin - 3.2, Math.max(bestY + 1.2, pageH - margin - 4.5));
  doc.setDrawColor(120);
  doc.setLineWidth(0.3);
  doc.line(margin, footerLineY, pageW - margin, footerLineY);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7.5);
  doc.setTextColor(80, 90, 100);
  doc.text('SL Lesson plan', pageW - margin, footerLineY + 3.2, { align: 'right' });
  if (d.date) {
    doc.text(String(d.date), margin, footerLineY + 3.2);
  }

  var fname = 'CBC_Lesson_Plan_' + String(d.lessonTitle || 'plan').replace(/[^\w]+/g, '_').slice(0, 40) + '.pdf';
  doc.save(fname);
}


function exportWord() {
  try { applySelfEvaluationSilent(); } catch (e) {}
  var content = document.getElementById('lessonPlanContent');
  if (!content || !content.innerHTML.trim()) { alert('Generate a lesson plan first.'); return; }
  var html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta charset="utf-8">'
    + '<style>body{font-family:Times New Roman,serif;font-size:11pt}table{border-collapse:collapse;width:100%}'
    + 'td,th{border:1px solid #000;padding:6px;vertical-align:top;font-size:11pt}.hdr,.lab{background:#d0d5dd;font-weight:bold}'
    + '.title{text-align:center;font-size:15pt;font-weight:bold;margin-bottom:10px}.plan-footer{text-align:right;margin-top:8px;font-size:8.5pt;font-weight:600;color:#475569}</style></head><body>'
    + content.innerHTML + '</body></html>';
  var blob = new Blob(['\ufeff', html], { type: 'application/msword' });
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'CBC_Lesson_Plan.doc';
  document.body.appendChild(a); a.click();
  setTimeout(function () { document.body.removeChild(a); }, 400);
}

function flatRows() {
  var d = LAST;
  if (!d) return [];
  return [
    ['Field', 'Value'],
    ['School Name', d.schoolName],
    ["Teacher's Name", d.teacherName],
    ['Term', d.term],
    ['Date', d.date],
    ['Subject', d.subject],
    ['Class', d.class],
    ['Unit No', d.unitNo],
    ['Lesson No', d.lessonNo + ' Out of ' + d.totalLessons],
    ['Duration', d.duration + ' min'],
    ['Class size', d.classSize],
    ['Special educational needs', d.specialNeeds],
    ['Unit title', d.unitTitle],
    ['Key unit competence', d.keyCompetence],
    ['Title of the lesson', d.lessonTitle],
    ['Instructional objectives', d.objective],
    ['Location', d.location],
    ['Learning materials', d.materials],
    ['References', d.references],
    ['Introduction (min)', d.introM],
    ["Introduction - Teacher", (d.introT || []).join(' | ')],
    ["Introduction - Learner", (d.introL || []).join(' | ')],
    ['Introduction - Competences', (d.introC || []).join(' | ')],
    ['Development (min)', d.devM],
    ["Development - Teacher", (d.devT || []).join(' | ')],
    ["Development - Learner", (d.devL || []).join(' | ')],
    ['Development - Competences', (d.devC || []).join(' | ')],
    ['Conclusion (min)', d.concM],
    ["Conclusion - Teacher", (d.concT || []).join(' | ')],
    ["Conclusion - Learner", (d.concL || []).join(' | ')],
    ['Conclusion - Competences', (d.concC || []).join(' | ')],
    [Lt.selfEval, d.selfEvaluation || '']
  ];
}

function applySelfEvaluationSilent() {
  if (!LAST) return;
  var se = document.getElementById('selfEvalInput');
  if (se) LAST.selfEvaluation = (se.value || '').trim();
  var cell = document.getElementById('selfEvalCell');
  if (cell) {
    var txt = LAST.selfEvaluation || '';
    cell.innerHTML = txt ? esc(txt).replace(/\n/g, '<br>') : '&nbsp;<br>&nbsp;<br>&nbsp;';
  }
}


function exportExcel() {
  try { applySelfEvaluationSilent(); } catch (e) {}
  if (!LAST) { alert('Generate a lesson plan first.'); return; }
  var ws = XLSX.utils.aoa_to_sheet(flatRows());
  ws['!cols'] = [{ wch: 32 }, { wch: 80 }];
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Lesson Plan');
  XLSX.writeFile(wb, 'CBC_Lesson_Plan.xlsx');
}

function exportCSV() {
  try { applySelfEvaluationSilent(); } catch (e) {}
  if (!LAST) { alert('Generate a lesson plan first.'); return; }
  var rows = flatRows();
  var csv = rows.map(function (r) {
    return r.map(function (c) {
      var s = String(c == null ? '' : c);
      if (/[",\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
      return s;
    }).join(',');
  }).join('\n');
  var blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8' });
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'CBC_Lesson_Plan.csv';
  document.body.appendChild(a); a.click();
  setTimeout(function () { document.body.removeChild(a); }, 400);
}


/* ========== AUTH SYSTEM (localStorage) ==========
   First registered user = Admin (approved automatically).
   Later users = pending until Admin approves.
================================================ */
var AUTH_USERS_KEY = 'cbc_auth_users_v1';
var AUTH_SESSION_KEY = 'cbc_auth_session_v1';

function authLoadUsers() {
  try {
    var raw = localStorage.getItem(AUTH_USERS_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch (e) { return []; }
}
function authSaveUsers(users) {
  localStorage.setItem(AUTH_USERS_KEY, JSON.stringify(users));
}
function authGetSession() {
  try {
    var raw = localStorage.getItem(AUTH_SESSION_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}
function authSetSession(user) {
  if (!user) {
    localStorage.removeItem(AUTH_SESSION_KEY);
    return;
  }
  localStorage.setItem(AUTH_SESSION_KEY, JSON.stringify({
    email: user.email,
    name: user.name,
    role: user.role,
    ts: Date.now()
  }));
}
function authFindUser(email) {
  var e = String(email || '').trim().toLowerCase();
  return authLoadUsers().find(function (u) { return u.email === e; }) || null;
}
function authCurrentUser() {
  var s = authGetSession();
  if (!s || !s.email) return null;
  var u = authFindUser(s.email);
  if (!u || u.status !== 'approved') {
    authSetSession(null);
    return null;
  }
  return u;
}

/* Simple password hash (SHA-256) */
function authHash(password) {
  return new Promise(function (resolve) {
    if (window.crypto && window.crypto.subtle) {
      var data = new TextEncoder().encode('cbc_salt_v1:' + password);
      window.crypto.subtle.digest('SHA-256', data).then(function (buf) {
        var arr = Array.from(new Uint8Array(buf));
        resolve(arr.map(function (b) { return b.toString(16).padStart(2, '0'); }).join(''));
      }).catch(function () {
        resolve(authFallbackHash(password));
      });
    } else {
      resolve(authFallbackHash(password));
    }
  });
}
function authFallbackHash(password) {
  var s = 'cbc_salt_v1:' + password;
  var h = 0;
  for (var i = 0; i < s.length; i++) {
    h = ((h << 5) - h) + s.charCodeAt(i);
    h |= 0;
  }
  // expand a bit
  var out = '';
  for (var j = 0; j < 8; j++) {
    h = ((h << 5) - h) + j * 31 + s.length;
    h |= 0;
    out += (h >>> 0).toString(16).padStart(8, '0');
  }
  return out;
}

function showAuthTab(tab) {
  var users = authLoadUsers();
  var hasAdmin = users.length > 0;
  var authSub = document.getElementById('authSub');
  if (authSub) {
    authSub.textContent = hasAdmin
      ? 'Sign in with the account created for you by the Administrator'
      : 'First time: register to become the Administrator';
  }
  var tabReg = document.getElementById('tabRegister');
  var regForm = document.getElementById('registerForm');
  var loginForm = document.getElementById('loginForm');
  var regSubmit = document.getElementById('regSubmitBtn');
  var regClosed = document.getElementById('regClosedMsg');
  var hint = document.getElementById('regHint');

  // If admin already exists, public registration is closed
  if (hasAdmin) {
    tabReg.style.display = 'none';
    if (tab === 'register') tab = 'login';
  } else {
    tabReg.style.display = '';
  }

  document.getElementById('tabLogin').classList.toggle('active', tab === 'login');
  tabReg.classList.toggle('active', tab === 'register');
  loginForm.style.display = tab === 'login' ? 'block' : 'none';
  regForm.style.display = tab === 'register' ? 'block' : 'none';
  clearAuthMsg();

  if (tab === 'register' && !hasAdmin) {
    if (regSubmit) regSubmit.style.display = '';
    if (regClosed) regClosed.style.display = 'none';
    // show all reg fields
    Array.prototype.forEach.call(regForm.querySelectorAll('.form-group'), function (g) { g.style.display = ''; });
    if (hint) {
      hint.style.display = 'block';
      hint.innerHTML = '<strong>First registration only:</strong> You will become the <strong>sole Administrator</strong>. After this, only you can create accounts for other teachers. No one else can register as admin.';
    }
  } else if (tab === 'register' && hasAdmin) {
    if (regSubmit) regSubmit.style.display = 'none';
    if (regClosed) regClosed.style.display = 'block';
    Array.prototype.forEach.call(regForm.querySelectorAll('.form-group'), function (g) { g.style.display = 'none'; });
    if (hint) hint.style.display = 'none';
  } else {
    if (hint) hint.style.display = 'none';
  }
}

function clearAuthMsg() {
  var m = document.getElementById('authMsg');
  m.className = 'auth-msg';
  m.textContent = '';
}
function setAuthMsg(text, type) {
  var m = document.getElementById('authMsg');
  m.className = 'auth-msg ' + (type || 'info');
  m.textContent = text;
}

function handleRegister(e) {
  e.preventDefault();
  clearAuthMsg();
  var users = authLoadUsers();
  if (users.length > 0) {
    setAuthMsg('Registration is closed. Only the Administrator can create teacher accounts.', 'error');
    showAuthTab('login');
    return false;
  }
  var name = document.getElementById('regName').value.trim();
  var email = document.getElementById('regEmail').value.trim().toLowerCase();
  var pass = document.getElementById('regPassword').value;
  var pass2 = document.getElementById('regPassword2').value;
  if (!name || !email || !pass) {
    setAuthMsg('Please fill all fields.', 'error');
    return false;
  }
  if (pass.length < 6) {
    setAuthMsg('Password must be at least 6 characters.', 'error');
    return false;
  }
  if (pass !== pass2) {
    setAuthMsg('Passwords do not match.', 'error');
    return false;
  }
  if (authFindUser(email)) {
    setAuthMsg('This email/username is already registered.', 'error');
    return false;
  }
  authHash(pass).then(function (hash) {
    var recovery = 'ADMIN-' + Math.random().toString(36).slice(2, 6).toUpperCase() + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
    var user = {
      id: 'u_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8),
      name: name,
      email: email,
      passwordHash: hash,
      passwordPlain: pass,
      role: 'admin',
      status: 'approved',
      createdAt: new Date().toISOString(),
      recoveryCode: recovery
    };
    authSaveUsers([user]);
    try { localStorage.setItem('cbc_admin_recovery_v1', recovery); } catch (e) {}
    authSetSession(user);
    setAuthMsg('Welcome! You are the sole Administrator.', 'ok');
    // Show recovery code once
    setTimeout(function () {
      alert('IMPORTANT — Save your Admin Recovery Code:\\n\\n' + recovery + '\\n\\nIf you forget your password, use this code on the login page to reset it.\\nWrite it down and keep it safe.');
      showApp();
    }, 400);
  });
  return false;
}

function handleLogin(e) {
  e.preventDefault();
  clearAuthMsg();
  var email = document.getElementById('loginEmail').value.trim().toLowerCase();
  var pass = document.getElementById('loginPassword').value;
  if (!email || !pass) {
    setAuthMsg('Enter email and password.', 'error');
    return false;
  }
  var user = authFindUser(email);
  if (!user) {
    setAuthMsg('No account found. Please register first.', 'error');
    return false;
  }
  authHash(pass).then(function (hash) {
    if (hash !== user.passwordHash) {
      setAuthMsg('Incorrect password.', 'error');
      return;
    }
    if (user.status === 'pending') {
      setAuthMsg('Your account is waiting for Administrator approval.', 'info');
      return;
    }
    if (user.status === 'rejected') {
      setAuthMsg('Your registration was rejected. Contact the Administrator.', 'error');
      return;
    }
    if (user.status === 'disabled') {
      setAuthMsg('Your account is disabled. Contact the Administrator.', 'error');
      return;
    }
    authSetSession(user);
    setAuthMsg('Login successful. Loading…', 'ok');
    setTimeout(function () { showApp(); }, 400);
  });
  return false;
}

function handleLogout() {
  authSetSession(null);
  showAuth();
}

function showAuth() {
  document.getElementById('authScreen').classList.add('active');
  document.getElementById('authScreen').style.display = 'block';
  document.getElementById('adminScreen').classList.remove('active');
  var app = document.getElementById('appScreen');
  if (app) { app.classList.remove('active'); app.style.display = 'none'; }
  clearAuthMsg();
  showAuthTab('login');
  try { document.getElementById('loginForm').reset(); } catch (e) {}
}

function showApp() {
  var user = authCurrentUser();
  if (!user || user.status !== 'approved') {
    showAuth();
    return;
  }
  document.getElementById('authScreen').classList.remove('active');
  document.getElementById('authScreen').style.display = 'none';
  document.getElementById('adminScreen').classList.remove('active');
  var app = document.getElementById('appScreen');
  app.classList.add('active');
  app.style.display = 'block';
  // user bar
  document.getElementById('userBarName').textContent = user.name + ' (' + user.email + ')';
  var roleEl = document.getElementById('userBarRole');
  roleEl.textContent = user.role === 'admin' ? 'ADMIN' : 'TEACHER';
  roleEl.className = 'badge' + (user.role === 'admin' ? '' : ' teacher');
  document.getElementById('btnAdminPanel').style.display = user.role === 'admin' ? 'inline-block' : 'none';
  // prefill teacher name if empty
  try {
    var tn = document.getElementById('teacherName');
    if (tn && !tn.value) tn.value = user.name;
  } catch (e) {}
}

function openAdmin() {
  var user = authCurrentUser();
  if (!user || user.role !== 'admin') {
    alert('Admin access only.');
    return;
  }
  document.getElementById('appScreen').classList.remove('active');
  document.getElementById('authScreen').classList.remove('active');
  document.getElementById('adminScreen').classList.add('active');
  renderAdminUsers();
}

function closeAdmin() {
  showApp();
}

function setAdminMsg(text, type) {
  var m = document.getElementById('adminMsg');
  m.className = 'auth-msg ' + (type || 'info');
  m.textContent = text || '';
  if (!text) m.className = 'auth-msg';
}

function renderAdminUsers() {
  var body = document.getElementById('adminUserBody');
  var users = authLoadUsers();
  var me = authCurrentUser();
  body.innerHTML = '';
  if (!users.length) {
    body.innerHTML = '<tr><td colspan="6">No users yet.</td></tr>';
    return;
  }
  users.sort(function (a, b) {
    var order = { pending: 0, approved: 1, rejected: 2, disabled: 3 };
    return (order[a.status] || 9) - (order[b.status] || 9);
  });
  users.forEach(function (u) {
    var tr = document.createElement('tr');
    var dateStr = '';
    try { dateStr = new Date(u.createdAt).toLocaleDateString(); } catch (e) { dateStr = '—'; }
    var actions = '<div class="admin-actions">';
    // View + Reset password for everyone including self
    actions += '<button type="button" class="enable" onclick="openViewPassword(\'' + u.id + '\')">View password</button>';
    actions += '<button type="button" class="approve" onclick="openResetPassword(\'' + u.id + '\')">Reset password</button>';
    if (u.email !== me.email && u.role !== 'admin') {
      if (u.status === 'pending') {
        actions += '<button type="button" class="approve" onclick="adminSetStatus(\'' + u.id + '\',\'approved\')">Approve</button>';
        actions += '<button type="button" class="reject" onclick="adminSetStatus(\'' + u.id + '\',\'rejected\')">Reject</button>';
      }
      if (u.status === 'approved') {
        actions += '<button type="button" class="disable" onclick="adminSetStatus(\'' + u.id + '\',\'disabled\')">Disable</button>';
      }
      if (u.status === 'disabled' || u.status === 'rejected') {
        actions += '<button type="button" class="enable" onclick="adminSetStatus(\'' + u.id + '\',\'approved\')">Enable</button>';
      }
      actions += '<button type="button" class="delete" onclick="adminDeleteUser(\'' + u.id + '\')">Delete</button>';
    } else if (u.email === me.email) {
      actions += ' <em style="font-size:.8rem;color:#64748b">You (Admin)</em>';
    }
    actions += '</div>';
    tr.innerHTML =
      '<td>' + escapeHtml(u.name) + '</td>' +
      '<td>' + escapeHtml(u.email) + '</td>' +
      '<td>' + (u.role === 'admin' ? 'Admin' : 'Teacher') + '</td>' +
      '<td><span class="status ' + u.status + '">' + u.status + '</span></td>' +
      '<td>' + dateStr + '</td>' +
      '<td>' + actions + '</td>';
    body.appendChild(tr);
  });
}

function escapeHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}


function adminCreateUser(e) {
  e.preventDefault();
  setAdminMsg('');
  var me = authCurrentUser();
  if (!me || me.role !== 'admin') {
    setAdminMsg('Only the Administrator can create users.', 'error');
    return false;
  }
  var name = document.getElementById('adminNewName').value.trim();
  var email = document.getElementById('adminNewEmail').value.trim().toLowerCase();
  var pass = document.getElementById('adminNewPass').value;
  var pass2 = document.getElementById('adminNewPass2').value;
  if (!name || !email || !pass) {
    setAdminMsg('Please fill all fields.', 'error');
    return false;
  }
  if (pass.length < 6) {
    setAdminMsg('Password must be at least 6 characters.', 'error');
    return false;
  }
  if (pass !== pass2) {
    setAdminMsg('Passwords do not match.', 'error');
    return false;
  }
  if (authFindUser(email)) {
    setAdminMsg('This email/username already exists.', 'error');
    return false;
  }
  authHash(pass).then(function (hash) {
    var users = authLoadUsers();
    var user = {
      id: 'u_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8),
      name: name,
      email: email,
      passwordHash: hash,
      passwordPlain: pass,
      role: 'teacher',
      status: 'approved',
      createdAt: new Date().toISOString(),
      createdBy: me.email
    };
    users.push(user);
    authSaveUsers(users);
    document.getElementById('adminCreateForm').reset();
    setAdminMsg('Teacher created: ' + email + ' | Password: ' + pass + ' — share these details with the teacher.', 'ok');
    renderAdminUsers();
  });
  return false;
}

function adminSetStatus(userId, status) {
  var me = authCurrentUser();
  if (!me || me.role !== 'admin') return;
  var users = authLoadUsers();
  var u = users.find(function (x) { return x.id === userId; });
  if (!u) return;
  if (u.role === 'admin') {
    setAdminMsg('Cannot change the Administrator account status.', 'error');
    return;
  }
  // Never allow promoting anyone to admin
  u.role = 'teacher';
  u.status = status;
  authSaveUsers(users);
  setAdminMsg('Updated: ' + u.email + ' → ' + status, 'ok');
  renderAdminUsers();
}

function adminDeleteUser(userId) {
  var me = authCurrentUser();
  if (!me || me.role !== 'admin') return;
  var users = authLoadUsers();
  var u = users.find(function (x) { return x.id === userId; });
  if (!u) return;
  if (u.role === 'admin' || u.email === me.email) {
    setAdminMsg('Cannot delete the Administrator.', 'error');
    return;
  }
  if (!confirm('Delete user ' + u.email + '? They will no longer be able to log in.')) return;
  users = users.filter(function (x) { return x.id !== userId; });
  authSaveUsers(users);
  setAdminMsg('Deleted: ' + u.email, 'ok');
  renderAdminUsers();
}


/* ===== Password view / reset / admin recovery ===== */
var _pwModalUserId = null;
var _pwModalMode = null; // 'view' | 'reset' | 'recovery'

function openViewPassword(userId) {
  var me = authCurrentUser();
  if (!me || me.role !== 'admin') { alert('Admin only.'); return; }
  var u = authLoadUsers().find(function (x) { return x.id === userId; });
  if (!u) return;
  _pwModalUserId = userId;
  _pwModalMode = 'view';
  document.getElementById('pwModalTitle').textContent = 'View password';
  document.getElementById('pwModalSub').textContent = u.name + ' (' + u.email + ')';
  document.getElementById('pwViewBox').style.display = 'block';
  document.getElementById('pwResetBox').style.display = 'none';
  document.getElementById('pwRecoveryBox').style.display = 'none';
  document.getElementById('pwBtnSave').style.display = 'none';
  var val = u.passwordPlain;
  if (!val) {
    val = '(Not available — password was set before this feature. Use Reset password to set a new one.)';
  }
  document.getElementById('pwViewValue').textContent = val;
  document.getElementById('pwModalBg').classList.add('open');
}

function openResetPassword(userId) {
  var me = authCurrentUser();
  if (!me || me.role !== 'admin') { alert('Admin only.'); return; }
  var u = authLoadUsers().find(function (x) { return x.id === userId; });
  if (!u) return;
  _pwModalUserId = userId;
  _pwModalMode = 'reset';
  document.getElementById('pwModalTitle').textContent = 'Reset password';
  document.getElementById('pwModalSub').textContent = u.name + ' (' + u.email + ') — set a new password';
  document.getElementById('pwViewBox').style.display = 'none';
  document.getElementById('pwResetBox').style.display = 'block';
  document.getElementById('pwRecoveryBox').style.display = 'none';
  document.getElementById('pwBtnSave').style.display = 'inline-block';
  document.getElementById('pwBtnSave').textContent = 'Save new password';
  document.getElementById('pwNewInput').value = '';
  document.getElementById('pwNewInput2').value = '';
  document.getElementById('pwModalBg').classList.add('open');
}

function openAdminRecovery() {
  _pwModalUserId = null;
  _pwModalMode = 'recovery';
  document.getElementById('pwModalTitle').textContent = 'Admin password recovery';
  document.getElementById('pwModalSub').textContent = 'Only for the Administrator who lost their password.';
  document.getElementById('pwViewBox').style.display = 'none';
  document.getElementById('pwResetBox').style.display = 'none';
  document.getElementById('pwRecoveryBox').style.display = 'block';
  document.getElementById('pwBtnSave').style.display = 'inline-block';
  document.getElementById('pwBtnSave').textContent = 'Reset admin password';
  document.getElementById('pwRecoveryCode').value = '';
  document.getElementById('pwAdminNew').value = '';
  document.getElementById('pwAdminNew2').value = '';
  document.getElementById('pwModalBg').classList.add('open');
}

function closePwModal() {
  document.getElementById('pwModalBg').classList.remove('open');
  _pwModalUserId = null;
  _pwModalMode = null;
}

function copyPwView() {
  var v = document.getElementById('pwViewValue').textContent;
  if (!v || v.indexOf('Not available') === 0) return;
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(v).then(function () { alert('Password copied.'); }).catch(function () { prompt('Copy this password:', v); });
  } else {
    prompt('Copy this password:', v);
  }
}

function savePwModal() {
  if (_pwModalMode === 'reset') {
    var pass = document.getElementById('pwNewInput').value;
    var pass2 = document.getElementById('pwNewInput2').value;
    if (!pass || pass.length < 6) { alert('Password must be at least 6 characters.'); return; }
    if (pass !== pass2) { alert('Passwords do not match.'); return; }
    var users = authLoadUsers();
    var u = users.find(function (x) { return x.id === _pwModalUserId; });
    if (!u) return;
    authHash(pass).then(function (hash) {
      u.passwordHash = hash;
      u.passwordPlain = pass;
      authSaveUsers(users);
      closePwModal();
      setAdminMsg('Password reset for ' + u.email + ' → new password: ' + pass, 'ok');
      renderAdminUsers();
      alert('Password updated for ' + u.email + '\\n\\nNew password: ' + pass + '\\n\\nShare it with the user.');
    });
    return;
  }
  if (_pwModalMode === 'recovery') {
    var code = document.getElementById('pwRecoveryCode').value.trim().toUpperCase();
    var pass = document.getElementById('pwAdminNew').value;
    var pass2 = document.getElementById('pwAdminNew2').value;
    if (!code) { alert('Enter recovery code.'); return; }
    if (!pass || pass.length < 6) { alert('Password must be at least 6 characters.'); return; }
    if (pass !== pass2) { alert('Passwords do not match.'); return; }
    var users = authLoadUsers();
    var admin = users.find(function (x) { return x.role === 'admin'; });
    if (!admin) { alert('No admin account found.'); return; }
    var stored = '';
    try { stored = (localStorage.getItem('cbc_admin_recovery_v1') || '').toUpperCase(); } catch (e) {}
    var ok = (admin.recoveryCode && admin.recoveryCode.toUpperCase() === code) || (stored && stored === code);
    if (!ok) {
      alert('Invalid recovery code.');
      return;
    }
    authHash(pass).then(function (hash) {
      admin.passwordHash = hash;
      admin.passwordPlain = pass;
      admin.status = 'approved';
      authSaveUsers(users);
      closePwModal();
      setAuthMsg('Admin password reset successfully. You can log in now.', 'ok');
      showAuthTab('login');
      alert('Admin password has been reset.\\n\\nNew password: ' + pass);
    });
  }
}



function showMyRecoveryCode() {
  var me = authCurrentUser();
  if (!me || me.role !== 'admin') return;
  var code = me.recoveryCode || '';
  try {
    if (!code) code = localStorage.getItem('cbc_admin_recovery_v1') || '';
  } catch (e) {}
  if (!code) {
    // generate and save if missing (upgrade path)
    code = 'ADMIN-' + Math.random().toString(36).slice(2, 6).toUpperCase() + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
    var users = authLoadUsers();
    var admin = users.find(function (x) { return x.id === me.id; });
    if (admin) {
      admin.recoveryCode = code;
      authSaveUsers(users);
      try { localStorage.setItem('cbc_admin_recovery_v1', code); } catch (e2) {}
    }
  }
  alert('Your Admin Recovery Code:\\n\\n' + code + '\\n\\nKeep it safe. Use it on the login page if you forget your password.');
}

function authInit() {
  // Always hide app first — never show lesson plan without login
  try {
    var app = document.getElementById('appScreen');
    var auth = document.getElementById('authScreen');
    var adm = document.getElementById('adminScreen');
    if (app) { app.classList.remove('active'); app.style.display = 'none'; }
    if (adm) { adm.classList.remove('active'); }
    if (auth) { auth.classList.add('active'); auth.style.display = 'block'; }
  } catch (e0) {}

  var user = authCurrentUser();
  if (user && user.status === 'approved') {
    showApp();
  } else {
    // Clear invalid session
    try { authSetSession(null); } catch (e1) {}
    showAuth();
  }
}

// Boot auth after DOM ready — always gate behind login
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', authInit);
} else {
  authInit();
}


try {
  var dateEl = document.getElementById('date');
  if (dateEl && !dateEl.value) {
    dateEl.value = new Date().toISOString().slice(0, 10);
  }
} catch (eDate) {}
</script>
</div><!-- /appScreen -->
</body>
</html>
