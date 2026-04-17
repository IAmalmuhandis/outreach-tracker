# YARE Outreach Assistant

Simple, responsive personal outreach web app for:
- Managing organizations and outreach status
- Storing contact lists per organization (email and phone)
- Logging calls, meetings, and follow-ups
- Importing initial organization list from `YARE_Abuja_Outreach_Master_List.docx`
- Keeping your outreach guide available inside the dashboard

## Run

1. Install dependencies:
   - `npm install`
2. Start:
   - `npm start`
3. Open:
   - `http://localhost:4000`

## Configuration

The app uses `.env`:
- `PORT=4000`
- `MONGODB_URI=<your mongodb connection string>`

## Notes

- The API uses MongoDB database name: `yareOutreach`.
- Existing guide file `yare_outreach_guide.html` is served at `/guide`.
- Use the "Import Organizations from DOCX" button on the top right to pull organizations from your master list.
