# Club Attendance

QR-code attendance for school clubs, built on Google Forms, Google Sheets and Apps Script.

## Why I built this

I'm the president of my school's debate club. This year we have 84 members, from prep-year students to 11th graders, and we meet every week.

Paper attendance doesn't hold up at that size. Calling 84 names eats a real part of a one-hour meeting, and anyone can sign a sheet for a friend who stayed home. I wanted something that takes a few seconds per student, can't be done from the couch, and still lets me fix things by hand when someone's phone is dead.

So we use one QR code. Students scan it, tap *Present*, and submit; their school Google account proves who they are. The spreadsheet does the rest: it marks who came, marks who didn't once the meeting is over, and keeps a running rate for every member.

One lesson along the way: I started with a list of 84 names and no email addresses. I guessed them from the school's username pattern, but a wrong guess means a student who showed up gets marked absent. That's why there's a registration form. Each student fills it in once, and their real address replaces the guess.

This repository is the system our club uses, with our members' data taken out, so other clubs can set it up too.

## What's in the repository

| File | What it is |
|---|---|
| [`Club_Attendance_Template.xlsx`](Club_Attendance_Template.xlsx) | Empty spreadsheet template with all sheets, formulas, dropdowns and colors. No data. |
| [`Code.gs`](Code.gs) | The Apps Script that runs the forms and fills the sheets. |

## Features

- **One QR code for the whole year.** The attendance form opens and closes automatically on the meeting times you enter.
- **Verified emails.** Google takes each address from the student's signed-in account, and responses from outside your school domain are rejected.
- **Automatic PRESENT / ABSENT** that you can override by hand. Manual edits are kept.
- **Optional meeting code.** Write a code for a meeting and the form asks for it. A photo of the QR code alone is then not enough to check in from home.
- **Registration form** that adds new members, fixes wrong email addresses and reactivates returning members. Suspicious cases are held for review.
- **Report and dashboard** with attendance per meeting, per student and overall.
- **No lost responses.** When a meeting ends, every form response is processed again, including any that a busy moment kept from being recorded.

## Requirements

- A Google account that can create Forms and Apps Script projects.
- For verified, school-only check-ins, a Google Workspace account (for example Google Workspace for Education). With a personal Google account, set `EMAIL_DOMAIN` to `''`. The forms then accept any Google account.

## Setup

1. Download **`Club_Attendance_Template.xlsx`** and upload it to Google Drive with the account that will run the club.
2. Open it with Google Sheets and choose **File → Save as Google Sheets**. Use only this Google Sheets copy from now on. The `.xlsx` file has no Extensions menu.
3. Under **File → Settings → Time zone**, choose your school's time zone. Meeting times are read in this time zone.
4. Open **Extensions → Apps Script**, delete the contents of `Code.gs` and paste in [`Code.gs`](Code.gs).
5. Edit `CONFIG` at the top of the script. At minimum, set `EMAIL_DOMAIN` (see [Configuration](#configuration)). Save.
6. Select **`setup`** in the function list and click **Run**. Authorize with the same account.
   - Setup creates both forms, installs three triggers and puts the links and QR codes on the **Dashboard**.
7. **Restrict both forms to your school.** The script cannot do this step. For each form (the edit links are on the Dashboard):
   - Click **Published → Responders → Manage**.
   - Under **Responder view**, choose your organization instead of *Anyone with the link*.
   - Click **Done**, then **Save**.
   - Under **Settings → Responses**, *Collect email addresses* should be **Verified** and *Limit to 1 response* should be **off**.
8. Reload the spreadsheet. An **Attendance** menu appears.
9. Test with a student account. In a private browser window, open the attendance form and check that it asks for a school sign-in.

Running `setup` again is safe. It reuses the same forms, QR codes and data, and refreshes the formulas. Run it again after pasting a new version of the script.

## Configuration

| Key | Default | Meaning |
|---|---|---|
| `CLUB_NAME` | `'Debate Club'` | Used in form titles and on the Dashboard. |
| `EMAIL_DOMAIN` | `'example.edu'` | Your school email domain, e.g. `'myschool.org'`. Responses from other domains are rejected. Use `''` to accept any Google account. Setup refuses to run while this is `'example.edu'`. |
| `LOCALE` | `'en-US'` | Used for upper/lower case in names and codes. Use `'tr-TR'` for Turkish (i/İ, ı/I). |
| `GRADES` | `[['9', 3], ['10', 2], ['11', 1], ['12', 0]]` | Grade choices on the registration form. The number is how many years after the current school year the student graduates; it fills *Graduation Year*. Add `['Prep', 4]` at the front for a prep year. |
| `ASK_SECTION` | `true` | Ask for a section (e.g. `B`). Class is stored as `10/B`. If `false`, class is just the grade. |
| `UPPERCASE_LAST_NAME` | `false` | Store last names in capitals (`Jane DOE`). |
| `SCHOOL_YEAR_START_MONTH` | `8` | The month the school year starts (8 = August). Used for graduation years. |
| `DEFAULT_DURATION_MIN` | `60` | Meeting length when *End* is left blank. |
| `OPEN_EARLY_MIN` | `5` | The form may open this many minutes before the start time. Responses in this window count. |
| `GRACE_MIN` | `5` | Responses this many minutes after the end still count. `ABSENT` is written after this. |
| `MAX_MEETINGS` | `40` | Number of meeting columns. The template is built for 40, so leave this unless you change the template too. |

## Sheets

| Sheet | You edit it? | Contents |
|---|---|---|
| **Dashboard** | No | Totals, form links, QR codes and the system **Status** line. |
| **Members** | Yes | One row per student: ID, name, class, email, `Active`/`Left`, `Unverified`/`Verified`, graduation year, note. |
| **Meetings** | Yes | One row per meeting: date, start, optional end, optional code, note. |
| **Attendance** | Yes (meeting cells) | One row per member, one column per meeting, plus Present / Absent / Rate. |
| **Report** | No | Present, absent and rate for each meeting. |
| **Responses** | No | Every attendance form response with its result. Rewritten after each meeting. |
| **Registrations** | No | Every registration form response with its result. |
| **Marks** | No, hidden | The script's own record of what it wrote. This is how it tells your manual edits apart. |

Don't rename the sheets or insert columns into them. Renaming the spreadsheet file itself is fine.

## Using it

### Members

Add students to **Members**, one per row. **ID** is required and must be unique, and **Email** must be the student's school address exactly. Add new students at the bottom. The script adds their Attendance row within five minutes, or right away if you use **Attendance → Check form now**.

When a student leaves, set their **Status** to `Left` instead of deleting the row. Their history is kept and they stop receiving `ABSENT`.

You can also skip typing members and let students use the registration form.

### Meetings

Write each meeting in the next empty row of **Meetings**:

| Date | Start | End | Code | Note |
|---|---|---|---|---|
| 2026-11-03 | 16:15 | | | |
| 2026-11-10 | 16:15 | 17:30 | Table | Guest speaker |

- **Start** accepts `16:15`, `16.15` or `4:15 PM`.
- **End** is optional. If blank, the meeting lasts `DEFAULT_DURATION_MIN` minutes.
- **Code** is optional. If set, the form asks for it during that meeting. Case, spaces and accents are ignored. Write the code before the meeting.
- You can sort the rows. Each meeting's Attendance column moves with it.
- To cancel a meeting, clear or delete its row. Its Attendance column is cleared.
- You can also enter a meeting that already happened. Within five minutes everyone is marked `ABSENT`; then set the students who came to `PRESENT`.

### During a meeting

- The script checks every five minutes, so the form opens up to five minutes after its scheduled time, sometimes a little later.
- The **Status** line on the Dashboard shows whether the form is open, the next meeting, or an error in a Meetings row.
- A valid response writes `PRESENT`. A second response from the same student is logged as `DUPLICATE`.

### After a meeting

After the end time plus `GRACE_MIN`, the script:

1. processes every form response again, recovering any that weren't recorded during a busy moment;
2. writes `ABSENT` for every active member who has no mark for that meeting.

### Manual changes

Any meeting cell in **Attendance** can be set to `PRESENT` or `ABSENT` by hand, for example for a student without a phone or one who left early. The script compares each cell with its hidden record in **Marks** and never overwrites a cell you changed.

### Wrong or missing emails

A response from an address that isn't in Members appears in **Responses** as `Email not found in Members`. That row shows the student's real address. Either:

- correct the email in **Members**, then run **Attendance → Reprocess responses**; or
- ask the student to fill in the registration form.

Either way, their earlier responses count from then on.

## Registration form

Students enter first name, last name, student ID, grade and (optionally) section. The email comes from their school account. Each submission is logged in **Registrations**:

| Result | What happened |
|---|---|
| `ADDED` | New student ID: added to Members as `Active` and `Verified`, with an Attendance row. |
| `UPDATED` | Existing student: email verified or corrected, `Left` changed back to `Active`, or class updated. |
| `REGISTERED` | Already registered; nothing changed. |
| `REVIEW` | **Not applied.** The email belongs to another student, the ID is linked to a different verified email, or the name doesn't match the listed name. Decide in Members by hand. |
| `REJECTED` | Not a school account, or invalid ID, name, grade or section. |

Rules for existing students:

- An `Unverified` email is replaced by the form's address only if the names match. The last name must match and the first name must appear in the listed name. Accents and middle names are ignored.
- Names already in Members are never changed.

If members were typed in by hand with guessed addresses, ask everyone to fill in the registration form once. This confirms or corrects every address.

Open or close registration at any time with **Attendance → Open / close registration**.

## Menu

| Item | What it does |
|---|---|
| Reprocess responses | Re-evaluates every attendance response against the current Members list. Manual edits are kept. |
| Check form now | Runs the five-minute check immediately. |
| Open / close registration | Toggles the registration form. |
| Run setup | Same as running `setup` from the editor. |

## Privacy

- All data stays in your own Google Drive: the spreadsheet and the two forms.
- The QR images are generated by [QuickChart](https://quickchart.io), which receives only the public form link.
- Share the spreadsheet only with club staff. It contains student names and emails.
- Don't commit a filled-in spreadsheet or member list to a public repository.

## Limitations

- Meetings must not overlap. One meeting per date works best.
- Up to 40 meetings per spreadsheet. Start a new copy of the template each school year.
- Forms open and close on a five-minute timer, not to the second.
- Students need to be signed in to their school account on their phone.
- Restricting responders to your organization is a manual step (setup step 7).

## Troubleshooting

| Problem | Fix |
|---|---|
| *Edit CONFIG at the top of the script…* | Set `EMAIL_DOMAIN`, save, run `setup` again. |
| *Missing sheet(s): …* | The script is attached to a spreadsheet that wasn't made from the template, or a sheet was renamed. Follow setup steps 1–2 and paste the script into that copy. |
| No **Extensions** menu | The file is still the `.xlsx`. Use **File → Save as Google Sheets** and open the copy. |
| Students see *You need permission* | They are signed in with the wrong account, or setup step 7 wasn't done for that form. |
| The form never opens | Check the Dashboard **Status** line. Check that **Triggers** (clock icon in Apps Script) lists `updateFormWindow`. Check the spreadsheet time zone. |
| Status shows `ERROR · Meetings row N…` | Fix the date or time in that row. Other meetings keep working. |
| Everyone is `ABSENT` for a meeting | Usually the meeting was entered after it ended, or the Members emails are wrong. Fix them and run **Reprocess responses**, or edit the cells. |
| Started over in a new copy | The old copy's triggers keep running. Delete them in the old Apps Script project (Triggers page), and delete the old forms. |

## Uninstalling

In Apps Script, open **Triggers** and delete the three triggers. Then delete the two forms from Drive.
