# Row limit update

Requested change: make long table displays show up to 5000 rows instead of 500/100 row caps.

Required changes in `public/index_replacement.html`:

- Replace every `.slice(0,500)` with `.slice(0,5000)`.
- Replace the Analytics table `.slice(0,100)` with `.slice(0,5000)`.

Affected areas:

- Invoice Control
- Data Quality
- Job Register
- Job Cards
- Analytics tables

Debtors & Collections already has pagination and is not affected by this note.
