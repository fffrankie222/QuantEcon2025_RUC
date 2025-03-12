from docx import Document
import openpyxl
import docxtpl
import docx2pdf

#读取表格
sheet = openpyxl.load_workbook('University_Major.xlsx').active


#替换并生成文档
for i in range(1,31):
    University_cell=sheet['A'+str(i+1)]
    Country_cell=sheet['E'+str(i+1)]
    for x in range(1,4):
        next_letter=chr(ord('A')+x)
        Major_cell=sheet[next_letter+str(i+1)]

        doc=docxtpl.DocxTemplate('文字文稿1.docx')
        content = {
            'University': University_cell.value,
            'Major': Major_cell.value,
            'Country': Country_cell.value
        }
        doc.render(content)

        output_name=University_cell.value+'_'+Major_cell.value+'.docx'
        doc.save(output_name)
        docx2pdf.convert(output_name)
        print('第'+str(i)+'个学校第'+str(x)+'个专业完成')
print('全部完成')
