import os
from tkinter import Tk, filedialog
from comtypes.client import CreateObject


def convert_presentation_to_pdf(ppt_file_path, pdf_file_path):
    """
    단일 PowerPoint 파일을 PDF로 변환
    """
    ppt_file_path = os.path.abspath(ppt_file_path)
    pdf_file_path = os.path.abspath(pdf_file_path)

    powerpoint_app = CreateObject("PowerPoint.Application")
    powerpoint_app.Visible = 1
    try:
        print(f"PowerPoint 파일 열기: {ppt_file_path}")
        presentation = powerpoint_app.Presentations.Open(ppt_file_path)
        presentation.SaveAs(pdf_file_path, 32)  # PDF로 저장
        print(f"PDF 저장 완료: {pdf_file_path}")
        presentation.Close()
    except Exception as error:
        print(f"파일 변환 실패: {ppt_file_path}, 오류: {error}")
    finally:
        powerpoint_app.Quit()


def convert_all_presentations_in_folder(folder_path):
    """
    선택한 폴더 내 모든 PowerPoint 파일을 PDF로 변환
    """
    output_folder = os.path.join(folder_path, "pdf변환")
    os.makedirs(output_folder, exist_ok=True)

    for file_name in os.listdir(folder_path):
        if file_name.lower().endswith((".pptx", ".ppt")):
            ppt_file_path = os.path.join(folder_path, file_name)
            pdf_file_name = os.path.splitext(file_name)[0] + ".pdf"
            pdf_file_path = os.path.join(output_folder, pdf_file_name)
            convert_presentation_to_pdf(ppt_file_path, pdf_file_path)


print("폴더를 선택하려면 1, 파일을 선택하려면 2를 입력하세요.")
choice = input("선택: ").strip()

tk_window = Tk()
tk_window.withdraw()

if choice == "1":
    selected_folder = filedialog.askdirectory(title="PPT 파일이 있는 폴더를 선택하세요")
    if not selected_folder:
        print("폴더가 선택되지 않았습니다.")
    else:
        print(f"선택된 폴더: {selected_folder}")
        convert_all_presentations_in_folder(selected_folder)
elif choice == "2":
    selected_file = filedialog.askopenfilename(
        title="PowerPoint 파일을 선택하세요",
        filetypes=[("PowerPoint Files", "*.pptx;*.ppt")]
    )
    if not selected_file:
        print("파일이 선택되지 않았습니다.")
    else:
        print(f"선택된 파일: {selected_file}")
        output_pdf_path = os.path.splitext(selected_file)[0] + ".pdf"
        convert_presentation_to_pdf(selected_file, output_pdf_path)
else:
    print("올바른 입력이 아닙니다. 프로그램을 종료합니다.")
