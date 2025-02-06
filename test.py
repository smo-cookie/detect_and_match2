import os
import zipfile
from tempfile import TemporaryDirectory
from lxml import etree
import re
import openai


def mask_text(text, regex_patterns):
    """
    텍스트에서 민감 정보를 마스킹
    :param text -> 입력파일의 텍스트
    :param regex_patterns -> 정규표현식 패턴
    :return -> 마스킹된 텍스트
    """
    if not text:
        return text
    for pattern in regex_patterns:
        text = re.sub(pattern, "****", text)
    return text


def process_xml_file(xml_path, regex_patterns):
    """
    XML 파일을 열어 텍스트 노드를 수정
    :param xml_path -> XML 파일 경로
    :param regex_patterns -> 개인정보 탐지 정규표현식 패턴 리스트
    """
    parser = etree.XMLParser(remove_blank_text=True)
    with open(xml_path, 'rb') as file:
        xml_tree = etree.parse(file, parser)

    # 모든 텍스트 노드(<w:t>)를 순회하며 마스킹
    for element in xml_tree.xpath("//w:t", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}):
        element.text = mask_text(element.text, regex_patterns)

    with open(xml_path, 'wb') as file:
        file.write(etree.tostring(xml_tree, pretty_print=True))


def process_comments_xml(comments_path, regex_patterns):
    """
    Word 문서의 comments.xml 파일에서 주석 텍스트를 마스킹
    :param comments_path -> comments.xml 파일 경로(주석에 존재하는 개인정보도 고려하기!)
    :param regex_patterns -> 개인정보 탐지 정규표현식 리스트
    """
    if os.path.exists(comments_path):
        process_xml_file(comments_path, regex_patterns)


def mask_sensitive_data_with_images(file_path, regex_patterns):
    """
    Word 문서에서 텍스트는 마스킹하고 이미지와 기타 요소는 유지
    :param file_path -> 원본 Word 파일 경로
    :param regex_patterns -> 개인정보를 탐지할 정규표현식 리스트
    :return -> 마스킹된 Word 파일 경로
    """
    # 임시 디렉토리 사용 -> zip 파일 압축 해제
    with TemporaryDirectory() as temp_dir:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # 텍스트 수정 (document.xml, comments.xml 등등..)
        document_xml_path = os.path.join(temp_dir, "word", "document.xml")
        if os.path.exists(document_xml_path):
            process_xml_file(document_xml_path, regex_patterns)

        # 주석 파일 처리(comments.xml)
        comments_xml_path = os.path.join(temp_dir, "word", "comments.xml")
        process_comments_xml(comments_xml_path, regex_patterns)

        # 수정된 파일을 다시 zip으로 압축하여 새로운 .docx 파일 생성
        new_file_path = file_path.replace(".docx", "(masked).docx")
        with zipfile.ZipFile(new_file_path, 'w') as zip_out:
            for foldername, subfolders, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zip_out.write(file_path, arcname)

    return new_file_path


if __name__ == "__main__":
    # 정규표현식 리스트 (전화번호, 이메일, 주민등록번호 등)
    patterns = [
        r'\b\d{2,3}-\d{3,4}-\d{4}\b',   # 전화번호
        r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b',  # 이메일
        r'\b\d{6}-\d{7}\b',  # 주민등록번호
    ]

    # 처리할 파일 경로받기
    input_file = input("마스킹할 Word 파일 경로를 입력하세요: ").strip()

    if not os.path.exists(input_file):
        print("파일이 존재하지 않습니다. 경로를 확인하세요.")
    else:
        # 개인정보 마스킹 파일 출력
        masked_file = mask_sensitive_data_with_images(input_file, patterns)
        print(f"마스킹된 파일이 저장되었습니다: {masked_file}")
