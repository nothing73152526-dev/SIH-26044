# Security Review — SkillBridge Prototype (SIH26044)

**Target reviewed:** `codestu.docx` — a Microsoft Word document containing the
HTML/CSS/JS source of the "SkillBridge" Smart India Hackathon prototype
(academia–industry collaboration portal).
**Review type:** Static (white-box) source review of the code embedded in the
repository. No live system was tested.
**Date:** 2026-10-28

> ### Scope note
> This review covers **only the code in this repository**. It is **not** a test
> of Bitbns or any third-party live website. Unauthenticated/offensive testing
> of a production system was explicitly *not* performed, as there is no verified
> authorization for it.

---

## Summary

The prototype is a **100% client-side static mockup**: a single HTML page with
inline CSS and JavaScript. It makes **no network requests**, has **no backend**,
and stores **no data**.

| Area | Result |
|------|--------|
| Hardcoded credentials / API keys / tokens | **None found** (verified by pattern scan) |
| Secrets in connection strings / JWTs / private keys | **None found** |
| Embedded URLs / endpoints | **None found** |
| Client-side injection sinks | **1 found** (DOM XSS, see F-1) |
| Access-control model | **Cosmetic only** — architectural risk for the real build (see F-2) |

Because there is no server, the live *impact* of the code-level issues is low
**today**. They matter because this prototype is the blueprint for a real,
multi-role portal that will handle student data, documents, and a "secure
document vault" (per the page footer). The findings below are written for that
next step.

---

## F-1 — DOM-based XSS in `postJob()`  (Severity: Medium when backend-connected)

**Location:** `postJob()` in the Industry view `<script>`.

The function concatenates three user-controlled `<input>` values directly into
`innerHTML`:

```js
function postJob(){
  const t=document.getElementById('jt').value||'New Role';
  const s=document.getElementById('jskills').value||'—';
  const d=document.getElementById('jdesc').value||'—';
  document.getElementById('postedJobs').innerHTML+=
   `<div class="job"><b>${t}</b><small>${d}</small><span class="chip">${s}</span>...`;
}
```

A recruiter entering a job title such as:

```
<img src=x onerror="alert(document.cookie)">
```

causes the payload to execute. In the current static demo this is only
**self-XSS** (the attacker and victim are the same person). **But** the moment
posted jobs are persisted and shown to *other* users (students browsing
listings — exactly what the prototype describes), this becomes **stored XSS**:
a malicious recruiter's job post runs script in every student's session.

### Fix
Build the node with `textContent` / `createElement` instead of HTML string
concatenation, so user input is never parsed as markup:

```js
function postJob(){
  const t=document.getElementById('jt').value||'New Role';
  const s=document.getElementById('jskills').value||'—';
  const d=document.getElementById('jdesc').value||'—';

  const job=document.createElement('div'); job.className='job';
  const title=document.createElement('b'); title.textContent=t;
  const desc=document.createElement('small'); desc.textContent=d;
  const chip=document.createElement('span'); chip.className='chip'; chip.textContent=s;
  const live=document.createElement('span');
  live.className='pill p-ok'; live.style.marginLeft='8px';
  live.textContent='Live ✓ · students notified';

  job.append(title, desc, chip, live);
  document.getElementById('postedJobs').append(job);

  document.getElementById('jt').value=
  document.getElementById('jskills').value=
  document.getElementById('jdesc').value='';
}
```

When the real backend exists, **also** escape/validate on the server and
re-encode on output — client-side fixes are not a substitute for server-side
output encoding.

---

## F-2 — Access control is cosmetic (Severity: High in the real build)

The role switcher (`setRole`) works purely by toggling a CSS `.hidden` class:

```js
function setRole(r, btn){
  ['student','industry','faculty','institution'].forEach(x=>{
    document.getElementById('view-'+x).classList.toggle('hidden', x!==r);
  });
}
```

Every role's view — including the **Institution Analytics** dashboard (student
counts, placement rates) and the **Industry recruiter** dashboard (candidate
lists, applications) — is present in the DOM for *every* visitor at all times.
"Switching role" just unhides a section; nothing is actually restricted.

The page footer advertises *"Role-based access · Secure document vault"*, but
none of that is enforced. As-is (static mock) there's no real data to leak.
In the production system this pattern is **Broken Access Control (OWASP A01)**
— the #1 web risk. A student could read recruiter/institution data simply by
un-hiding a `<div>` or calling the API directly.

### Fix (for the real build)
- Enforce authorization **server-side** on every API endpoint and every
  document fetch, keyed to the authenticated user's role — never by what the
  browser chooses to render.
- Send each client only the data its role is entitled to; don't ship all roles'
  data to the browser and hide it with CSS.
- Protect the "document vault" with per-object authorization checks (verify the
  requester owns/may access each document), not just role checks.

---

## F-3 — Defense-in-depth notes (low priority / for the production build)

1. **`<base target="_blank">` without `rel`.** All links open new tabs. When
   real external links are added, include `rel="noopener noreferrer"` to
   prevent reverse tabnabbing (the opened page gaining `window.opener` access).
2. **No Content-Security-Policy.** A CSP (`default-src 'self'`, no inline
   script once code is externalized) would have blunted F-1 as a second layer.
   Note the current inline `<script>`/`<style>` and `onclick=` handlers would
   need to move to external files for a strict CSP.
3. **No input validation/length limits** on the recruiter form fields.
4. **Mock data only.** The names, companies, and figures in the page
   (Aarav, AyurData Labs, stipends, counts) are demonstration data, not real
   PII or secrets — confirmed, nothing sensitive is committed.

---

## What was checked and found clean

- Hardcoded passwords, API keys, secrets, bearer tokens — **none**.
- AWS keys, private keys, JWTs, DB connection strings — **none**.
- Outbound URLs / hardcoded endpoints — **none**.
- Any form of network/data exfiltration in the client code — **none**.

## Recommended next steps
1. Apply the F-1 fix in the prototype now (cheap, prevents the pattern carrying
   forward).
2. Treat F-2 as a hard requirement for the real build: authorization is
   server-side, always.
3. Add CSP + `rel="noopener"` and server-side validation when the backend lands.
