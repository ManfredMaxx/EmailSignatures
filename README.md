# Signify — Email Signature Manager for Microsoft Outlook

Signify adds a personal email signature panel directly inside Outlook. Each person on your team sets up their own signature once, inserts it with a single click when sending email, and cannot accidentally send without it — Outlook will warn them if they try.

---

## How it works in plain English

- A **Signify** button appears in the ribbon at the top of every compose window
- Clicking it opens a sidebar where you fill in your name, title, phone, compliance text, and anything else your signature needs
- When you're ready to send, click **Insert into this email** — your signature drops in right above any quoted reply thread, exactly where it belongs
- If you click Send without inserting your signature, Outlook stops you and reminds you

Signatures are saved privately inside your own Microsoft account. Nobody else can see or edit them.

---

## Part 1 — Administrator Setup (done once, before handing off to users)

> This section is for whoever is setting up Signify for the team. End users can skip to Part 2.

### Step 1 — Enable GitHub Pages

The Signify files are hosted on GitHub. You need to turn on GitHub Pages so Outlook can reach them.

1. Go to **https://github.com/ManfredMaxx/EmailSignatures**
2. Click the **Settings** tab near the top of the page
3. In the left sidebar, click **Pages**
4. Under **Source**, click the dropdown that says "None" and select **Deploy from a branch**
5. Set the branch to **main** and the folder to **/ (root)**
6. Click **Save**
7. Wait 2–3 minutes. Refresh the page. You should see a green banner that says *"Your site is live at https://manfredmaxx.github.io/EmailSignatures/"*
8. Click that link to confirm the page loads. You should see a plain XML file.

If you don't see the green banner after 5 minutes, try navigating away and coming back.

---

### Step 2 — Install Signify for your users in Microsoft 365

1. Open a web browser and go to **https://admin.microsoft.com**
2. Sign in with your **Microsoft 365 admin account** (the one with admin privileges, not your everyday email)
3. In the left navigation, click **Settings**, then click **Integrated apps**
4. Click **Upload custom apps** (button in the upper area of the page)
5. Under "How would you like to upload this app?", select **Provide link to manifest file**
6. In the URL box, paste exactly this address:
   ```
   https://manfredmaxx.github.io/EmailSignatures/manifest.xml
   ```
7. Click **Validate**. You should see "Signify" appear with a green checkmark.
8. Click **Next**
9. On the "Add Users" screen, select **Specific users/groups**, then search for and add each person who should have Signify
10. Click **Next**, review, and click **Finish**

Microsoft 365 will now install Signify for the selected users. This can take anywhere from **30 minutes to 24 hours** to appear in their Outlook. In practice it's usually under 2 hours.

---

### Step 3 — Let your users know

Once deployed, send your users a message along these lines:

> *"A tool called Signify has been added to your Outlook. It lets you set up and insert your email signature. Look for the Signify button in the ribbon when you open a new email. Follow the setup instructions to get started."*

Then point them to Part 2 of this guide.

---

## Part 2 — User Guide

> Everything below is for end users. No technical knowledge required.

---

### Finding Signify in Outlook

Signify appears as a button in the **ribbon** — the toolbar that runs along the top of the compose window.

**To find it:**
1. Open Outlook (the desktop app, the web version at outlook.com, or your company's Outlook web address)
2. Click **New Email** or reply to any existing email
3. Look along the top toolbar for a button labelled **Signify**

> If you don't see the Signify button, try closing and reopening Outlook. If it still doesn't appear after 24 hours of being set up by your administrator, contact them.

---

### Setting up your signature for the first time

You only need to do this once. After that, your signature is saved and ready every time.

1. Open a new email (click **New Email**)
2. Click the **Signify** button in the toolbar — a panel opens on the right side of the screen
3. Fill in your details:

   | Field | What to put here |
   |---|---|
   | Full Name | Your first and last name |
   | Job Title | Your official title (e.g. "Licensed Real Estate Broker") |
   | Company | Your company or firm name |
   | Department | Your team or division, if applicable |
   | Office Phone | Your direct office number |
   | Mobile | Your mobile number, if you want to include it |
   | Email Address | Your email address |
   | Website | Your company website (e.g. "company.com") |
   | LinkedIn | Your LinkedIn profile URL, if desired |
   | Office Address | Your mailing address |
   | Accent Color | The color used for the signature border and links — pick your brand color |
   | **Compliance Text** | **License numbers, regulatory language, certifications** — this appears in a highlighted box in your signature |
   | Legal Disclaimer | Fine-print text that appears at the very bottom in small grey text |

4. Click the **Preview** tab at the top to see exactly how your signature will look
5. When you're happy with it, click **Save**

Your signature is now saved. You do not need to fill this in again unless you want to make changes.

---

### Sending an email with your signature

Every time you write a new email or reply:

1. Write your message as normal
2. When you're ready to send, click the **Signify** button
3. Click **Insert into this email**
4. Your signature appears immediately below your message, above any quoted thread
5. Click **Send**

That's it.

---

### What happens if you forget to insert your signature

If you click **Send** without inserting your signature, Outlook will **stop the email from sending** and show you a message:

> *"No signature found. Click Signify in the ribbon, then click Insert into this email before sending."*

You will see two options:
- **Cancel** — goes back to your email so you can add the signature and send properly
- **Send Anyway** — overrides the warning and sends without a signature (use only in genuine emergencies)

In almost all cases you should click **Cancel**, open Signify, insert your signature, and then send.

---

### Updating your signature

If your title changes, you get a new phone number, or you need to update your compliance text:

1. Open any new email or reply
2. Click the **Signify** button
3. The panel opens with your previously saved details already filled in
4. Edit whichever fields have changed
5. Click **Save**
6. Click **Insert into this email** to use the updated version in the current email

Your new signature is saved immediately and will be used in all future emails.

---

### Resetting your signature (starting over)

If you want to completely clear your signature and start fresh:

1. Open Signify
2. Click the **Reset** button at the bottom right
3. The button label changes to **Tap again to confirm** — click it once more to confirm
4. All your saved details are cleared

You can then fill in new details and save as normal.

---

## Troubleshooting

**Signify button doesn't appear in Outlook**
- Close Outlook completely and reopen it
- If using Outlook on the web, try refreshing the page (F5)
- If it still doesn't appear, wait up to 24 hours after your administrator set it up — Microsoft 365 can take time to push add-ins to all users

**My signature disappeared after I updated Outlook**
- Open Signify — your saved details should still be there
- Click **Insert into this email** to re-insert

**The Insert button says "No active email found"**
- Signify must be used from inside a compose or reply window
- Make sure you have a compose window open before clicking Insert

**I can see the Signify panel but clicking Save shows an error**
- Check your internet connection
- Try closing and reopening the compose window, then open Signify again

**My compliance text or disclaimer isn't showing**
- Make sure the Compliance Text and Legal Disclaimer fields are filled in and you've clicked **Save** after editing
- Click **Preview** to confirm the text appears before inserting

---

## Quick reference card

| What you want to do | How to do it |
|---|---|
| Set up your signature | Open new email → click Signify → fill in fields → Save |
| Insert your signature | Click Signify → Insert into this email |
| Preview your signature | Click Signify → click Preview tab |
| Update your signature | Click Signify → edit fields → Save |
| Start over | Click Signify → Reset → Tap again to confirm |
