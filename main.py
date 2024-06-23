from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse
import pandas as pd
from io import BytesIO

app = FastAPI()


@app.post("/upload-excel/")
async def upload_excel(file: UploadFile = File(...)):
    if not file.filename.endswith('.xlsx'):
        raise HTTPException(status_code=400, detail="File type must be .xlsx")

    try:
        # Read the Excel file into a DataFrame
        content = await file.read()
        df = pd.read_excel(BytesIO(content))

        # Data processing
        return_df = pd.DataFrame(index=df.index)
        return_df['שם פרטי'] = df['שם פרטי']
        return_df['שם משפחה'] = df['שם משפחה']
        return_df['Email'] = ''
        return_df['טלפון'] = ''
        return_df[['Email', 'טלפון']] = df.apply(identify_and_assign, axis=1)[
            ['Email', 'טלפון']]
        return_df['כתובת-רחוב'] = ''
        return_df['כתובת-עיר'] = ''
        return_df['כתובת מדינה'] = ''
        return_df['כתובת האתר'] = df['כתובת האתר']
        return_df['סטאטוס'] = ''
        return_df['מלל חופשי'] = ''
        return_df['סוג בקשה מהאינטרנט'] = ''
        return_df['מוכרן'] = ''
        return_df[['תאריך', 'שעה']] = df['פתיחת קריאה'].astype(
            str).str.split(' ', expand=True)
        return_df['קוד משתמש מהאתר'] = df['קוד משתמש מהאתר']
        return_df['SITE_URL'] = 'GLASSIX'
        return_df['קוד מקור הלידה'] = ''
        return_df['מותג'] = ''
        return_df['מוצר'] = ''

        # Save the modified DataFrame to an Excel file in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            return_df.to_excel(writer, index=False, sheet_name='Sheet1')

        output.seek(0)

        # Return the Excel file to the client
        headers = {
            "Content-Disposition": "attachment; filename=modified_file.xlsx"
        }
        return StreamingResponse(output, headers=headers, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


def identify_and_assign(row):
    contact_info = str(row['מזהה לקוח'])
    if '@' in contact_info:
        row['Email'] = contact_info
        row['טלפון'] = ''
    else:
        row['Email'] = ''
        row['טלפון'] = contact_info
    return row
