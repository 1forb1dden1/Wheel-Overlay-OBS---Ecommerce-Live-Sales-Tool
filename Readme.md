# Energy Break — Wheel controls (step by step)

This readme only describes **buttons and on-screen labels for the prize wheel**. The moving strip lives in **Spin & controls** and in **Open HTML wheel** (same thing you can put in OBS).

https://github.com/user-attachments/assets/15d4f8f4-c794-40f9-a24b-23190030d04d

## Quick reference — what each wheel-related button does

| Button | What it does |
|--------|----------------|
| **SPIN** | Spin once; prize **committed** when done (unless dry run). Use for fill-skip flows. |
| **SUPER SPIN** | Spin; then choose **KEEP** or one **REROLL** → then **KEEP**. |
| **REROLL** | After SUPER only — spin again once. |
| **KEEP** | After SUPER only — confirm and save landed prize. |
| **Skip spot** | No spin; inserts empty prize row / advances sequence. |
| **Undo spin** | Undo last eligible spin/skip-like action. |
| **Fill a skipped spot…** | Arms a skipped row; **next SPIN** fills it. |
| **Cancel fill mode** | Leave that mode without spinning. |
| **Open HTML wheel…** | Browser copy of **OBS** strip + labels. |

---

**1 — What Main window (overlay)**  
<img width="465" height="282" alt="image" src="https://github.com/user-attachments/assets/e9502d1a-e5e1-4b33-89a6-fff1e7487c9a" />

1. Turn on **Show setup (list path, dry run, log)**.  
2. Set your prize list path (or **Browse…**). The wheel reads SKUs and quantities from this file so the system knows where to draw from. 
3. Optionally turn on **Dry run (PRACTICE MODE)** — Practice only, this does not actually update or write anything onto the winners sheet.  

**2 — Open the wheel panel**  

1. Click **Spin & controls** on the overlay, this is typically opened by default.  
<img width="810" height="1069" alt="image" src="https://github.com/user-attachments/assets/4f5808a0-c5f6-4cfd-a79c-13d0c5629c2b" />

## Labels on and around the wheel (Spin & controls)

**3 — Normal draw: SPIN**  
Use Case: Regular Spin for a customer when their payment goes through and you need to do a wheel spin for them.

1. Click **SPIN**.  
2. The landed prize is **saved** and updates the winners list.xlsx file (unless dry run) is on.  
3. **Undo spin** may turn on afterward (see Edit spin). You might have to use this if a payment fail for a buyer or if you accidently did a spin that you want to reverse. 

**4 — Super draw: SUPER SPIN (1x REROLL OPTION)**  
Use Case: Entertain and hype up the shows, get the customers to have a chance to spin twice if they don't like their initial pull.

1. Click **SUPER SPIN**.  
2. Wait for the strip to stop.  
3. Two extra buttons appear: **REROLL** and **KEEP**.  
4. **KEEP** — keep this result; it is saved like a normal prize (respects **Dry run**).  
5. **REROLL** — spin **one** more time; then you must click **KEEP** on the **final** result to finish.  

## Edit spin section

**5 — Skip spot**  
Use Case: If a customer's payment fails for a certain spot, just skip it. If you need to re-fill it later just use "filled a skip spot". 

1. Click **Skip spot**.  
2. Confirm if asked.  


**6 — Undo spin**  
Use Case: If you accidently did a spin or if the customers payment failed afterwards. You can undo the spin. 

1. Enabled only when something can be undone and the wheel is idle.  
2. Click **Undo spin**.  
3. Confirm if asked.  
4. Rolls back the last eligible action (for example last completed spin or skip), so the wheel state and files match again.

---

## Fill skipped spot section

**7 — Attach a prize to an old empty skip**  
Use Case: If the customer's payment initally failed and you had to skip their spot. But they fixed their payment and now it has processed. You can go back and re-fill the spot for them via a spin.
Note: If you need to do a re-spin for them, just undo the spin and spin once again as super spin does not work for re-fills. 

1. Click **Fill a skipped spot**.  
2. Choose a spot in the dialog; confirm **Use this spot for next SPIN** (wording matches the UI).  
3. The spot line switches to fill-skip wording.  


**8 — Cancel if you change your mind**  
Use Case: If you accidently clicked fill mode and you need to cancel it, use this button.

1. Click **Cancel fill mode**.  
2. You return to normal “next spot” behavior **without** using that skipped row.

## Same wheel in a browser (OBS)

**9 — Open HTML wheel**  

1. In **Spin & controls**, under **HTML wheel (OBS)**, click **Open HTML wheel…**.  
2. Your browser opens a page showing the **same strip**, spot/next-draw text, and **Total prizes left** as the HTML feed.  

**10 — OBS (optional)**  

1. In OBS, add a **Browser** source with URL **`http://127.0.0.1:8765/`** (fixed port every run).
2. Start Energy Break before streaming — the app serves that page on localhost. Check the event log if the port was busy (a different port may be used).
