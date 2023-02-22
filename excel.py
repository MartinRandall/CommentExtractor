import xlsxwriter

from extract import Paragraph
 
def write_to_excel(filename, comments: list[Paragraph] ):
    excel_filename = filename.replace('.docx', '-comments.xlsx')
    workbook = xlsxwriter.Workbook(excel_filename)
    
    worksheet = workbook.add_worksheet('Comments')
    
    worksheet.freeze_panes(1, 0)
    dateFormat = workbook.add_format({'num_format': 'dd/mm/yy hh:mm:ss'})
    dateFormat.set_align('top')
    
    title_format = workbook.add_format()
    title_format.set_bold()
    
    default_format = workbook.add_format()
    default_format.set_align('top')
    
    centre_format = workbook.add_format()
    centre_format.set_center_across()
    centre_format.set_align('top')
    
    wrap_format = workbook.add_format()
    wrap_format.set_text_wrap()
    wrap_format.set_align('top')
    
    worksheet.set_column('A:A', 30)
    worksheet.set_column('B:C', 20)
    worksheet.set_column('E:E', 30)
    worksheet.set_column('F:F', 100)
    
    worksheet.write('A1', 'Comment', title_format)
    worksheet.write('B1', 'Author', title_format)
    worksheet.write('C1', 'Date', title_format)
    worksheet.write('D1', 'Initials', title_format)
    worksheet.write('E1', 'Highlight Text', title_format)
    worksheet.write('F1', 'Paragraph', title_format)
    
    line = 2
    for paragraph in comments:
      for comment in paragraph.comments:
        worksheet.write(f'A{line}', comment.text, wrap_format)    
        worksheet.write(f'B{line}', comment.author, default_format)
        worksheet.write(f'C{line}', comment.date, dateFormat)
        worksheet.write(f'D{line}', comment.initials, centre_format)
        worksheet.write(f'E{line}', comment.highlightText, wrap_format)
        worksheet.write(f'F{line}', paragraph.text, wrap_format)
        line += 1
        
    workbook.close()