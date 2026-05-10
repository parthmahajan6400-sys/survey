<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>Household Demographic Survey</title>
<link href="https://fonts.googleapis.com/css2?family=Sora:wght@400;500;600;700&family=DM+Sans:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root {
    --c1: #FF6B6B;
    --c2: #FF8E53;
    --c3: #FFC837;
    --c4: #11998e;
    --c5: #38ef7d;
    --c6: #667eea;
    --c7: #764ba2;
    --bg: #0f0c29;
    --bg2: #1a1640;
    --card: rgba(255,255,255,0.07);
    --card-border: rgba(255,255,255,0.13);
    --text: #ffffff;
    --text2: rgba(255,255,255,0.65);
    --text3: rgba(255,255,255,0.38);
    --radius: 20px;
    --radius-sm: 12px;
  }
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  html { scroll-behavior: smooth; }
  body {
    font-family: 'DM Sans', sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
    overflow-x: hidden;
  }

  /* Animated background */
  body::before {
    content: '';
    position: fixed; inset: 0; z-index: 0;
    background:
      radial-gradient(ellipse at 15% 30%, rgba(102,126,234,0.28) 0%, transparent 55%),
      radial-gradient(ellipse at 85% 70%, rgba(255,107,107,0.22) 0%, transparent 55%),
      radial-gradient(ellipse at 50% 10%, rgba(56,239,125,0.12) 0%, transparent 50%);
    animation: bgshift 12s ease-in-out infinite alternate;
  }
  @keyframes bgshift {
    0%   { opacity: 1; transform: scale(1); }
    100% { opacity: .8; transform: scale(1.05); }
  }

  .wrap { position: relative; z-index: 1; min-height: 100vh; }

  /* ── HEADER ── */
  .header {
    padding: 20px 20px 0;
    display: flex; justify-content: space-between; align-items: center;
  }
  .brand {
    display: flex; align-items: center; gap: 10px;
  }
  .brand-dot {
    width: 32px; height: 32px; border-radius: 10px;
    background: linear-gradient(135deg, var(--c6), var(--c1));
    display: flex; align-items: center; justify-content: center;
    font-size: 16px;
  }
  .brand-name {
    font-family: 'Sora', sans-serif;
    font-size: 14px; font-weight: 600;
    background: linear-gradient(90deg, var(--c6), var(--c1));
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    background-clip: text;
  }
  .admin-toggle {
    font-size: 12px; color: var(--text3);
    background: var(--card); border: 1px solid var(--card-border);
    padding: 7px 14px; border-radius: 999px; cursor: pointer;
    transition: all .2s; font-family: 'DM Sans', sans-serif;
  }
  .admin-toggle:hover { color: var(--text2); border-color: rgba(255,255,255,0.25); }

  /* ── PAGES ── */
  .page { display: none; padding: 24px 20px 48px; max-width: 540px; margin: 0 auto; }
  .page.active { display: block; animation: fadeUp .4s ease; }
  @keyframes fadeUp {
    from { opacity: 0; transform: translateY(18px); }
    to   { opacity: 1; transform: translateY(0); }
  }

  /* ── HERO ── */
  .hero { text-align: center; padding: 32px 0 28px; }
  .hero-icon {
    width: 72px; height: 72px; border-radius: 22px; margin: 0 auto 18px;
    background: linear-gradient(135deg, #667eea, #764ba2);
    display: flex; align-items: center; justify-content: center;
    font-size: 32px;
    box-shadow: 0 8px 32px rgba(102,126,234,0.4);
  }
  .hero h1 {
    font-family: 'Sora', sans-serif;
    font-size: 26px; font-weight: 700; line-height: 1.25;
    margin-bottom: 10px;
  }
  .hero p { font-size: 14px; color: var(--text2); line-height: 1.65; }

  /* ── PROGRESS ── */
  .progress-wrap { margin-bottom: 24px; }
  .progress-label {
    display: flex; justify-content: space-between;
    font-size: 12px; color: var(--text3); margin-bottom: 8px;
  }
  .progress-track {
    height: 5px; background: rgba(255,255,255,0.1);
    border-radius: 999px; overflow: hidden;
  }
  .progress-fill {
    height: 100%; border-radius: 999px;
    background: linear-gradient(90deg, var(--c6), var(--c1));
    transition: width .5s cubic-bezier(.4,0,.2,1);
  }

  /* ── QUESTION CARD ── */
  .q-card {
    background: var(--card);
    border: 1px solid var(--card-border);
    border-radius: var(--radius);
    padding: 22px 18px;
    margin-bottom: 14px;
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
  }
  .q-tag {
    display: inline-flex; align-items: center; gap: 6px;
    font-size: 10px; font-weight: 600; letter-spacing: .1em; text-transform: uppercase;
    color: #fff; padding: 4px 10px; border-radius: 999px; margin-bottom: 12px;
  }
  .q-tag.t1 { background: linear-gradient(90deg,#667eea,#764ba2); }
  .q-tag.t2 { background: linear-gradient(90deg,#FF6B6B,#FF8E53); }
  .q-tag.t3 { background: linear-gradient(90deg,#11998e,#38ef7d); }
  .q-tag.t4 { background: linear-gradient(90deg,#FFC837,#FF8008); }
  .q-text {
    font-family: 'Sora', sans-serif;
    font-size: 16px; font-weight: 600; line-height: 1.45;
    margin-bottom: 16px;
  }
  .options { display: grid; gap: 10px; }
  .opt-label {
    display: flex; align-items: center; gap: 12px;
    background: rgba(255,255,255,0.05);
    border: 1.5px solid rgba(255,255,255,0.1);
    border-radius: var(--radius-sm);
    padding: 13px 15px; cursor: pointer;
    transition: all .2s; font-size: 14px; color: var(--text2);
    -webkit-tap-highlight-color: transparent;
    user-select: none;
  }
  .opt-label:hover { background: rgba(255,255,255,0.1); border-color: rgba(255,255,255,0.22); color: var(--text); }
  .opt-label input[type=radio] { display: none; }
  .radio-dot {
    width: 20px; height: 20px; border-radius: 50%; flex-shrink: 0;
    border: 2px solid rgba(255,255,255,0.25);
    background: transparent; transition: all .2s;
    display: flex; align-items: center; justify-content: center;
  }
  .opt-label.selected {
    background: rgba(102,126,234,0.18);
    border-color: var(--c6);
    color: #fff;
  }
  .opt-label.selected .radio-dot {
    background: linear-gradient(135deg, var(--c6), var(--c1));
    border-color: transparent;
  }
  .opt-label.selected .radio-dot::after {
    content: ''; width: 7px; height: 7px; border-radius: 50%; background: #fff;
  }

  /* Color variants for selected states per question type */
  .q-card.c2 .opt-label.selected { background: rgba(255,107,107,0.15); border-color: var(--c1); }
  .q-card.c2 .opt-label.selected .radio-dot { background: linear-gradient(135deg,var(--c1),var(--c2)); }
  .q-card.c3 .opt-label.selected { background: rgba(17,153,142,0.18); border-color: var(--c4); }
  .q-card.c3 .opt-label.selected .radio-dot { background: linear-gradient(135deg,var(--c4),var(--c5)); }
  .q-card.c4 .opt-label.selected { background: rgba(255,200,55,0.13); border-color: var(--c3); }
  .q-card.c4 .opt-label.selected .radio-dot { background: linear-gradient(135deg,var(--c3),var(--c2)); }

  /* ── SUBMIT BUTTON ── */
  .submit-btn {
    width: 100%; padding: 16px;
    background: linear-gradient(135deg, var(--c6), var(--c1));
    color: #fff; border: none; border-radius: var(--radius-sm);
    font-family: 'Sora', sans-serif;
    font-size: 16px; font-weight: 600; cursor: pointer;
    margin-top: 8px;
    transition: opacity .2s, transform .15s;
    box-shadow: 0 6px 24px rgba(102,126,234,0.35);
    letter-spacing: .02em;
  }
  .submit-btn:hover:not(:disabled) { opacity: .9; transform: translateY(-1px); }
  .submit-btn:active:not(:disabled) { transform: scale(.98); }
  .submit-btn:disabled { opacity: .35; cursor: not-allowed; box-shadow: none; }

  /* ── SUCCESS PAGE ── */
  .success-center { text-align: center; padding: 48px 0; }
  .success-anim {
    width: 80px; height: 80px; border-radius: 50%; margin: 0 auto 24px;
    background: linear-gradient(135deg, var(--c4), var(--c5));
    display: flex; align-items: center; justify-content: center;
    font-size: 38px;
    box-shadow: 0 8px 32px rgba(17,153,142,.4);
    animation: popIn .5s cubic-bezier(.175,.885,.32,1.275);
  }
  @keyframes popIn {
    from { transform: scale(0); opacity: 0; }
    to   { transform: scale(1); opacity: 1; }
  }
  .success-center h2 {
    font-family: 'Sora', sans-serif;
    font-size: 28px; font-weight: 700; margin-bottom: 10px;
  }
  .success-center p { font-size: 15px; color: var(--text2); line-height: 1.6; margin-bottom: 32px; }
  .ghost-btn {
    display: inline-block; padding: 13px 28px;
    background: var(--card); border: 1px solid var(--card-border);
    color: var(--text); border-radius: var(--radius-sm);
    font-family: 'Sora', sans-serif; font-size: 14px; font-weight: 600;
    cursor: pointer; transition: all .2s;
  }
  .ghost-btn:hover { background: rgba(255,255,255,0.12); }

  /* ── LOGIN PAGE ── */
  .login-card {
    background: var(--card); border: 1px solid var(--card-border);
    border-radius: var(--radius); padding: 30px 24px;
    backdrop-filter: blur(16px); -webkit-backdrop-filter: blur(16px);
    margin-top: 16px;
  }
  .login-icon {
    width: 52px; height: 52px; border-radius: 16px; margin: 0 auto 18px;
    background: linear-gradient(135deg, #667eea, #764ba2);
    display: flex; align-items: center; justify-content: center;
    font-size: 24px;
  }
  .login-card h2 {
    font-family: 'Sora', sans-serif;
    font-size: 20px; font-weight: 700; text-align: center; margin-bottom: 6px;
  }
  .login-card .sub { font-size: 13px; color: var(--text2); text-align: center; margin-bottom: 24px; }
  .field { margin-bottom: 14px; }
  .field label {
    display: block; font-size: 11px; font-weight: 600;
    letter-spacing: .08em; text-transform: uppercase;
    color: var(--text3); margin-bottom: 7px;
  }
  .field input {
    width: 100%; padding: 13px 15px;
    background: rgba(255,255,255,0.07);
    border: 1.5px solid rgba(255,255,255,0.12);
    border-radius: var(--radius-sm);
    color: #fff; font-size: 15px; font-family: 'DM Sans', sans-serif;
    outline: none; transition: border-color .2s;
  }
  .field input:focus { border-color: var(--c6); }
  .field input::placeholder { color: var(--text3); }
  .err-msg { font-size: 13px; color: #FF6B6B; margin-top: 4px; text-align: center; min-height: 18px; }
  .back-link {
    display: block; text-align: center; margin-top: 16px;
    font-size: 13px; color: var(--text3); cursor: pointer; background: none; border: none;
    font-family: 'DM Sans', sans-serif;
  }
  .back-link:hover { color: var(--text2); }

  /* ── ADMIN DASHBOARD ── */
  .admin-header {
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 20px; flex-wrap: wrap; gap: 12px;
  }
  .admin-title {
    font-family: 'Sora', sans-serif;
    font-size: 22px; font-weight: 700;
  }
  .admin-actions { display: flex; gap: 8px; }
  .export-btn {
    padding: 10px 18px;
    background: linear-gradient(135deg, var(--c4), var(--c5));
    color: #fff; border: none; border-radius: var(--radius-sm);
    font-family: 'Sora', sans-serif;
    font-size: 13px; font-weight: 600; cursor: pointer;
    box-shadow: 0 4px 16px rgba(17,153,142,.3);
    transition: opacity .2s;
  }
  .export-btn:hover { opacity: .88; }
  .logout-btn {
    padding: 10px 16px;
    background: rgba(255,255,255,0.07); border: 1px solid var(--card-border);
    color: var(--text2); border-radius: var(--radius-sm);
    font-family: 'DM Sans', sans-serif; font-size: 13px; cursor: pointer;
    transition: all .2s;
  }
  .logout-btn:hover { background: rgba(255,255,255,.12); color: #fff; }
  .stats-grid {
    display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; margin-bottom: 20px;
  }
  .stat-card {
    background: var(--card); border: 1px solid var(--card-border);
    border-radius: var(--radius-sm); padding: 16px;
    text-align: center;
  }
  .stat-val { font-family: 'Sora', sans-serif; font-size: 32px; font-weight: 700; }
  .stat-val.grad1 { background: linear-gradient(90deg,var(--c6),var(--c1)); -webkit-background-clip:text; -webkit-text-fill-color:transparent; background-clip:text; }
  .stat-val.grad2 { background: linear-gradient(90deg,var(--c1),var(--c2)); -webkit-background-clip:text; -webkit-text-fill-color:transparent; background-clip:text; }
  .stat-val.grad3 { background: linear-gradient(90deg,var(--c4),var(--c5)); -webkit-background-clip:text; -webkit-text-fill-color:transparent; background-clip:text; }
  .stat-val.grad4 { background: linear-gradient(90deg,var(--c3),var(--c2)); -webkit-background-clip:text; -webkit-text-fill-color:transparent; background-clip:text; }
  .stat-lbl { font-size: 12px; color: var(--text3); margin-top: 4px; }
  .table-scroll { overflow-x: auto; border-radius: var(--radius-sm); }
  table { width: 100%; border-collapse: collapse; font-size: 13px; min-width: 700px; }
  thead th {
    padding: 11px 14px; text-align: left;
    font-size: 10px; font-weight: 600; letter-spacing: .08em; text-transform: uppercase;
    color: var(--text3);
    background: rgba(255,255,255,0.05);
    border-bottom: 1px solid var(--card-border);
  }
  tbody td {
    padding: 12px 14px;
    border-bottom: 1px solid rgba(255,255,255,0.05);
    color: var(--text2); vertical-align: middle;
  }
  tbody tr:last-child td { border-bottom: none; }
  tbody tr:hover td { background: rgba(255,255,255,0.04); color: var(--text); }
  .badge {
    display: inline-block; padding: 3px 10px; border-radius: 999px;
    font-size: 11px; font-weight: 600; white-space: nowrap;
  }
  .badge-male   { background: rgba(102,126,234,.25); color: #a5b4fc; }
  .badge-female { background: rgba(255,107,107,.22); color: #fca5a5; }
  .badge-other  { background: rgba(255,255,255,.1);  color: var(--text2); }
  .no-data {
    text-align: center; padding: 48px 24px;
    color: var(--text3); font-size: 15px;
    background: var(--card); border-radius: var(--radius-sm);
  }
  .no-data span { display: block; font-size: 40px; margin-bottom: 12px; }

  /* Clear data */
  .danger-btn {
    padding: 10px 16px;
    background: rgba(255,107,107,0.12); border: 1px solid rgba(255,107,107,0.3);
    color: #FF6B6B; border-radius: var(--radius-sm);
    font-family: 'DM Sans', sans-serif; font-size: 13px; cursor: pointer;
    transition: all .2s;
  }
  .danger-btn:hover { background: rgba(255,107,107,.22); }

  /* Scrollbar */
  ::-webkit-scrollbar { width: 5px; height: 5px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: rgba(255,255,255,.15); border-radius: 999px; }
</style>
</head>
<body>
<div class="wrap">

  <!-- HEADER -->
  <div class="header">
    <div class="brand">
      <div class="brand-dot">📋</div>
      <div class="brand-name">DemoSurvey</div>
    </div>
    <button class="admin-toggle" id="adminToggleBtn">🔐 Admin</button>
  </div>

  <!-- ══════════ SURVEY PAGE ══════════ -->
  <div class="page active" id="surveyPage">
    <div class="hero">
      <div class="hero-icon">📊</div>
      <h1>Household Demographic Survey</h1>
      <p>Help us understand your community better. All responses are confidential and used for research only.</p>
    </div>

    <div class="progress-wrap">
      <div class="progress-label">
        <span id="progressText">0 of 8 answered</span>
        <span id="progressPct">0%</span>
      </div>
      <div class="progress-track">
        <div class="progress-fill" id="progressFill" style="width:0%"></div>
      </div>
    </div>

    <!-- Q1 Age -->
    <div class="q-card c1">
      <div class="q-tag t1">01 · Demographics</div>
      <div class="q-text">What is your age group?</div>
      <div class="options" id="opts-age">
        <label class="opt-label"><input type="radio" name="age" value="Below 18"><div class="radio-dot"></div>Below 18</label>
        <label class="opt-label"><input type="radio" name="age" value="18 – 35"><div class="radio-dot"></div>18 – 35</label>
        <label class="opt-label"><input type="radio" name="age" value="36 – 60"><div class="radio-dot"></div>36 – 60</label>
        <label class="opt-label"><input type="radio" name="age" value="Above 60"><div class="radio-dot"></div>Above 60</label>
      </div>
    </div>

    <!-- Q2 Sex -->
    <div class="q-card c2">
      <div class="q-tag t2">02 · Demographics</div>
      <div class="q-text">What is your sex?</div>
      <div class="options" id="opts-sex">
        <label class="opt-label"><input type="radio" name="sex" value="Male"><div class="radio-dot"></div>Male</label>
        <label class="opt-label"><input type="radio" name="sex" value="Female"><div class="radio-dot"></div>Female</label>
        <label class="opt-label"><input type="radio" name="sex" value="Transgender"><div class="radio-dot"></div>Transgender</label>
        <label class="opt-label"><input type="radio" name="sex" value="Prefer not to say"><div class="radio-dot"></div>Prefer not to say</label>
      </div>
    </div>

    <!-- Q3 Education -->
    <div class="q-card c3">
      <div class="q-tag t3">03 · Demographics</div>
      <div class="q-text">What is your highest level of education?</div>
      <div class="options" id="opts-education">
        <label class="opt-label"><input type="radio" name="education" value="No formal education"><div class="radio-dot"></div>No formal education</label>
        <label class="opt-label"><input type="radio" name="education" value="Primary (up to class 5)"><div class="radio-dot"></div>Primary (up to class 5)</label>
        <label class="opt-label"><input type="radio" name="education" value="Secondary (class 6–12)"><div class="radio-dot"></div>Secondary (class 6–12)</label>
        <label class="opt-label"><input type="radio" name="education" value="Graduate & above"><div class="radio-dot"></div>Graduate & above</label>
      </div>
    </div>

    <!-- Q4 Marital -->
    <div class="q-card c4">
      <div class="q-tag t4">04 · Demographics</div>
      <div class="q-text">What is your marital status?</div>
      <div class="options" id="opts-marital">
        <label class="opt-label"><input type="radio" name="marital" value="Unmarried"><div class="radio-dot"></div>Unmarried</label>
        <label class="opt-label"><input type="radio" name="marital" value="Married"><div class="radio-dot"></div>Married</label>
        <label class="opt-label"><input type="radio" name="marital" value="Widowed"><div class="radio-dot"></div>Widowed</label>
        <label class="opt-label"><input type="radio" name="marital" value="Divorced / Separated"><div class="radio-dot"></div>Divorced / Separated</label>
      </div>
    </div>

    <!-- Q5 Household -->
    <div class="q-card c1">
      <div class="q-tag t1">05 · Household</div>
      <div class="q-text">What is the size of your household?</div>
      <div class="options" id="opts-household">
        <label class="opt-label"><input type="radio" name="household" value="1 – 2 members"><div class="radio-dot"></div>1 – 2 members</label>
        <label class="opt-label"><input type="radio" name="household" value="3 – 4 members"><div class="radio-dot"></div>3 – 4 members</label>
        <label class="opt-label"><input type="radio" name="household" value="5 – 6 members"><div class="radio-dot"></div>5 – 6 members</label>
        <label class="opt-label"><input type="radio" name="household" value="7 or more"><div class="radio-dot"></div>7 or more members</label>
      </div>
    </div>

    <!-- Q6 Income -->
    <div class="q-card c2">
      <div class="q-tag t2">06 · Household</div>
      <div class="q-text">What is your main source of income?</div>
      <div class="options" id="opts-income">
        <label class="opt-label"><input type="radio" name="income" value="Agriculture / Farming"><div class="radio-dot"></div>Agriculture / Farming</label>
        <label class="opt-label"><input type="radio" name="income" value="Wage / Daily labour"><div class="radio-dot"></div>Wage / Daily labour</label>
        <label class="opt-label"><input type="radio" name="income" value="Business / Self-employment"><div class="radio-dot"></div>Business / Self-employment</label>
        <label class="opt-label"><input type="radio" name="income" value="Government job / Pension"><div class="radio-dot"></div>Government job / Pension</label>
      </div>
    </div>

    <!-- Q7 Landholding -->
    <div class="q-card c3">
      <div class="q-tag t3">07 · Household</div>
      <div class="q-text">How much agricultural land does your household own?</div>
      <div class="options" id="opts-land">
        <label class="opt-label"><input type="radio" name="land" value="No land (landless)"><div class="radio-dot"></div>No land (landless)</label>
        <label class="opt-label"><input type="radio" name="land" value="Less than 1 acre"><div class="radio-dot"></div>Less than 1 acre</label>
        <label class="opt-label"><input type="radio" name="land" value="1 – 3 acres"><div class="radio-dot"></div>1 – 3 acres</label>
        <label class="opt-label"><input type="radio" name="land" value="More than 3 acres"><div class="radio-dot"></div>More than 3 acres</label>
      </div>
    </div>

    <!-- Q8 SHG -->
    <div class="q-card c4">
      <div class="q-tag t4">08 · SHG Membership</div>
      <div class="q-text">Are you a member of a Self-Help Group (SHG)?</div>
      <div class="options" id="opts-shg">
        <label class="opt-label"><input type="radio" name="shg" value="Yes, active member"><div class="radio-dot"></div>Yes, active member</label>
        <label class="opt-label"><input type="radio" name="shg" value="Yes, inactive member"><div class="radio-dot"></div>Yes, inactive member</label>
        <label class="opt-label"><input type="radio" name="shg" value="No, but interested"><div class="radio-dot"></div>No, but interested</label>
        <label class="opt-label"><input type="radio" name="shg" value="No, not interested"><div class="radio-dot"></div>No, not interested</label>
      </div>
    </div>

    <button class="submit-btn" id="submitBtn" disabled>Submit Responses →</button>
  </div>

  <!-- ══════════ SUCCESS PAGE ══════════ -->
  <div class="page" id="successPage">
    <div class="success-center">
      <div class="success-anim">✅</div>
      <h2>Response Recorded!</h2>
      <p>Thank you for taking the time to complete this survey. Your input is valuable for our research.</p>
      <button class="ghost-btn" id="anotherBtn">Submit another response</button>
    </div>
  </div>

  <!-- ══════════ LOGIN PAGE ══════════ -->
  <div class="page" id="loginPage">
    <div class="login-card">
      <div class="login-icon">🔐</div>
      <h2>Admin Login</h2>
      <p class="sub">Sign in to view and export all survey responses.</p>
      <div class="field">
        <label>Username</label>
        <input type="text" id="loginUser" placeholder="Enter username" autocomplete="username">
      </div>
      <div class="field">
        <label>Password</label>
        <input type="password" id="loginPass" placeholder="Enter password" autocomplete="current-password">
      </div>
      <div class="err-msg" id="loginErr"></div>
      <button class="submit-btn" id="loginBtn" style="margin-top:8px">Sign In →</button>
      <button class="back-link" id="backToSurveyBtn">← Back to survey</button>
    </div>
  </div>

  <!-- ══════════ ADMIN DASHBOARD PAGE ══════════ -->
  <div class="page" id="adminPage">
    <div class="admin-header">
      <div class="admin-title">📊 Responses</div>
      <div class="admin-actions">
        <button class="export-btn" id="exportBtn">⬇ Export Excel</button>
        <button class="danger-btn" id="clearBtn">🗑 Clear</button>
        <button class="logout-btn" id="logoutBtn">Sign out</button>
      </div>
    </div>
    <div class="stats-grid">
      <div class="stat-card"><div class="stat-val grad1" id="statTotal">0</div><div class="stat-lbl">Total responses</div></div>
      <div class="stat-card"><div class="stat-val grad2" id="statFemale">0</div><div class="stat-lbl">Female respondents</div></div>
      <div class="stat-card"><div class="stat-val grad3" id="statSHG">0</div><div class="stat-lbl">SHG members</div></div>
      <div class="stat-card"><div class="stat-val grad4" id="statMale">0</div><div class="stat-lbl">Male respondents</div></div>
    </div>
    <div class="table-scroll">
      <div id="tableWrap"></div>
    </div>
  </div>

</div><!-- /wrap -->

<script>
  const STORAGE_KEY = 'demosurvey_v1';
  const FIELDS = ['age','sex','education','marital','household','income','land','shg'];
  let answers = {};

  // ── Storage helpers ──
  function load() {
    try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]'); }
    catch { return []; }
  }
  function save(arr) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(arr));
  }

  // ── Page switching ──
  function showPage(id) {
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    document.getElementById(id).classList.add('active');
    window.scrollTo(0, 0);
  }

  // ── Progress ──
  function updateProgress() {
    const done = FIELDS.filter(f => answers[f]).length;
    const total = FIELDS.length;
    const pct = Math.round((done / total) * 100);
    document.getElementById('progressFill').style.width = pct + '%';
    document.getElementById('progressText').textContent = done + ' of ' + total + ' answered';
    document.getElementById('progressPct').textContent = pct + '%';
    document.getElementById('submitBtn').disabled = done < total;
  }

  // ── Radio selection ──
  document.querySelectorAll('.opt-label').forEach(label => {
    label.addEventListener('click', function() {
      const input = this.querySelector('input[type=radio]');
      const name = input.name;
      // Deselect siblings
      document.querySelectorAll(`input[name="${name}"]`).forEach(inp => {
        inp.closest('.opt-label').classList.remove('selected');
      });
      input.checked = true;
      this.classList.add('selected');
      answers[name] = input.value;
      updateProgress();
    });
  });

  // ── Submit ──
  document.getElementById('submitBtn').addEventListener('click', function() {
    if (FIELDS.some(f => !answers[f])) return;
    const record = { ...answers, timestamp: new Date().toISOString() };
    const arr = load();
    arr.push(record);
    save(arr);
    showPage('successPage');
  });

  // ── Another response ──
  document.getElementById('anotherBtn').addEventListener('click', function() {
    answers = {};
    document.querySelectorAll('.opt-label').forEach(l => l.classList.remove('selected'));
    document.querySelectorAll('input[type=radio]').forEach(i => i.checked = false);
    updateProgress();
    showPage('surveyPage');
  });

  // ── Admin toggle ──
  document.getElementById('adminToggleBtn').addEventListener('click', function() {
    showPage('loginPage');
    document.getElementById('loginErr').textContent = '';
    document.getElementById('loginUser').value = '';
    document.getElementById('loginPass').value = '';
  });

  // ── Login ──
  function doLogin() {
    const u = document.getElementById('loginUser').value.trim();
    const p = document.getElementById('loginPass').value;
    if (u === 'admin' && p === 'admin') {
      document.getElementById('loginErr').textContent = '';
      renderAdmin();
      showPage('adminPage');
    } else {
      document.getElementById('loginErr').textContent = '✗ Incorrect username or password.';
    }
  }
  document.getElementById('loginBtn').addEventListener('click', doLogin);
  document.getElementById('loginPass').addEventListener('keydown', function(e) {
    if (e.key === 'Enter') doLogin();
  });

  // ── Back to survey ──
  document.getElementById('backToSurveyBtn').addEventListener('click', function() {
    showPage('surveyPage');
  });

  // ── Logout ──
  document.getElementById('logoutBtn').addEventListener('click', function() {
    showPage('surveyPage');
  });

  // ── Clear data ──
  document.getElementById('clearBtn').addEventListener('click', function() {
    if (confirm('Are you sure you want to delete ALL survey responses? This cannot be undone.')) {
      save([]);
      renderAdmin();
    }
  });

  // ── Render admin table ──
  function renderAdmin() {
    const data = load();
    document.getElementById('statTotal').textContent = data.length;
    document.getElementById('statMale').textContent = data.filter(r => r.sex === 'Male').length;
    document.getElementById('statFemale').textContent = data.filter(r => r.sex === 'Female').length;
    document.getElementById('statSHG').textContent = data.filter(r => r.shg && r.shg.startsWith('Yes')).length;

    const wrap = document.getElementById('tableWrap');
    if (data.length === 0) {
      wrap.innerHTML = '<div class="no-data"><span>📭</span>No responses yet.<br>Share this file to start collecting data.</div>';
      return;
    }

    const cols = ['#','Timestamp','Age','Sex','Education','Marital Status','Household Size','Income Source','Landholding','SHG Membership'];
    let thead = '<thead><tr>' + cols.map(c => `<th>${c}</th>`).join('') + '</tr></thead>';
    let tbody = '<tbody>' + data.map((r, i) => {
      const sexClass = r.sex === 'Male' ? 'badge-male' : r.sex === 'Female' ? 'badge-female' : 'badge-other';
      const ts = r.timestamp ? new Date(r.timestamp).toLocaleString() : '—';
      return `<tr>
        <td>${i + 1}</td>
        <td style="white-space:nowrap;font-size:12px">${ts}</td>
        <td>${r.age || '—'}</td>
        <td><span class="badge ${sexClass}">${r.sex || '—'}</span></td>
        <td>${r.education || '—'}</td>
        <td>${r.marital || '—'}</td>
        <td>${r.household || '—'}</td>
        <td>${r.income || '—'}</td>
        <td>${r.land || '—'}</td>
        <td>${r.shg || '—'}</td>
      </tr>`;
    }).join('') + '</tbody>';

    wrap.innerHTML = `<table>${thead}${tbody}</table>`;
  }

  // ── Export Excel ──
  document.getElementById('exportBtn').addEventListener('click', function() {
    const data = load();
    if (data.length === 0) { alert('No responses to export yet!'); return; }
    const headers = ['#','Timestamp','Age Group','Sex','Education','Marital Status','Household Size','Income Source','Landholding','SHG Membership'];
    const rows = data.map((r, i) => [
      i + 1,
      r.timestamp ? new Date(r.timestamp).toLocaleString() : '',
      r.age || '', r.sex || '', r.education || '', r.marital || '',
      r.household || '', r.income || '', r.land || '', r.shg || ''
    ]);
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws['!cols'] = headers.map(() => ({ wch: 24 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Survey Responses');
    const date = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `survey_responses_${date}.xlsx`);
  });

  updateProgress();
</script>
</body>
</html>
