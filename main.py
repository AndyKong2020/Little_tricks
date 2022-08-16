import comtypes.client
import os

director = r' '
FileList = map(lambda x: director + '\\' + x, os.listdir(director))

for file in FileList:
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    slides = powerpoint.Presentations.Open(file)
    # 保存PDF
    slides.SaveAs(f"{file.split('.')[0]}.pdf", 32)
    slides.Close()
    print("完成" + f"{file.split('.')[0]}.pdf")