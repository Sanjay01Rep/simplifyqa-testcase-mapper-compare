# ICEA LION Testcase Review — How to unzip, set up, and run

Windows guide for managers and reviewers who received **`ICEA-Lion-Testcase-Review-share.zip`**.

See **[START-HERE.md](../START-HERE.md)** (same content, at the ZIP root).

Use this checklist when sharing the tool for review. No code changes required.

## 1. Access

- Repo: https://github.com/Sanjay01Rep/simplifyqa-testcase-mapper-compare
- Local folder must include `Client doc/`, `Kenya doc/` (and optional Kenya original) — these are **not** in git.

## 2. Run

```bash
npm install
npm start
```

Open http://localhost:3100

## 3. Demo scenarios (suggested)

| # | Scenario | Expected |
|---|----------|----------|
| 1 | Map one AP / GL / Investment client file | SimplifyQA Excel + log download |
| 2 | Compare Gen UG vs Life UG for one module | Common + Unique A + Unique B sheets |
| 3 | Pick a `STEPS_DIFFER` case from the log | Explain why it stayed off Common (extra step / different text) |
| 4 | `npm test` | All checks green |

## 4. What reviewers should not expect

- Live Preview / static HTML alone will not run APIs.
- Soft matching ignores typos; it does **not** ignore different business content (KRA vs URA, extra steps, different dates when they change meaning).

## 5. Attachments for the review email

- Link to GitHub repo
- Short screen recording or 3–4 screenshots (Map, Compare, output Excel tabs, log)
- One sample anonymized output Excel (if client data policy allows)
