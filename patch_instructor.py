from docx import Document

path = r'c:\Users\Owner\Documents\itec5040-assignment-2\Week2_Lab_Shruti_Malik.docx'
doc = Document(path)

changed = 0
for para in doc.paragraphs:
    for run in para.runs:
        if '[Instructor Name]' in run.text:
            run.text = run.text.replace('[Instructor Name]', 'Dr. C')
            changed += 1

doc.save(path)
print(f'Done - replaced {changed} instance(s).')
