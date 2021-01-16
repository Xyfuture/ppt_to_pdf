import comtypes.client
import os
import shutil

def mkdir_ppt():
    cwd = os.getcwd()
    new_dir = cwd+'\\ppt'
    if not os.path.exists(new_dir):
        os.makedirs(new_dir)

def move_file(file_name):
    shutil.move(file_name,"ppt");

def init_powerpoint():
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    return powerpoint

def ppt_to_pdf(powerpoint, inputFileName, outputFileName, formatType = 32):
    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName.replace('.pptx','')
        outputFileName = outputFileName.replace('.ppt','')
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
    deck.Close()

def convert_files_in_folder(powerpoint, folder):
    files = os.listdir(folder)
    pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]
    for pptfile in pptfiles:
        fullpath = os.path.join(cwd, pptfile)
        ppt_to_pdf(powerpoint, fullpath, fullpath)
        move_file(pptfile)


if __name__ == "__main__":
    powerpoint = init_powerpoint()
    cwd = os.getcwd()
    mkdir_ppt()
    convert_files_in_folder(powerpoint, cwd)
    powerpoint.Quit()