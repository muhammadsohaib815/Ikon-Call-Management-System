# IKON Call Performance System

The application uses a vanilla JavaScript frontend with a Python FastAPI backend. Pandas handles record data, openpyxl preserves and edits Excel workbooks, and SQLAlchemy connects to PostgreSQL.

## Run locally

1. Keep the existing database settings in `.env`.
2. Install dependencies with `pip install -r requirements.txt`.
3. Start the app with `python -m uvicorn main:app --reload --port 3000`.
4. Open `http://localhost:3000`.

## Project structure

- `main.py` contains the FastAPI backend and API routes.
- `public/index.html`, `public/styles.css`, and `public/app.js` are the single active frontend files served by FastAPI.
- `data/` contains reference/import files; dashboard data is read from PostgreSQL tables.
- `.env`, `requirements.txt`, and `sql.txt` contain runtime settings, Python dependencies, and the database schema reference.

The dashboard reads data from `public.ff_master_staff`, `public.call_attendence`, and `public.live_call_details`. CSV/XLSX files are supported only as optional uploads that are validated and saved into the database; the dashboard does not read the `data/csv/` folder.

Older phase/checkpoint copies and the unused Node.js placeholder have been removed so there is only one frontend version to maintain.

## Authentication and permissions

The application uses PostgreSQL tables `public.app_users` and `public.app_sessions` for login sessions and permissions. Passwords are stored as PBKDF2-SHA256 hashes, never as plain text.

Set the first administrator in the local `.env` file:

```env
AUTH_ADMIN_USERNAME=admin
AUTH_ADMIN_PASSWORD=change-this-password
```

On first login, open **User Management** to create additional users. A user can be given read, add, edit, and delete permissions. The backend enforces these permissions on every API request, so hiding buttons in the frontend is only a usability improvement—not the security control.
