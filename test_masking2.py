import os
import json
import re
import openai
import zipfile
from tempfile import TemporaryDirectory
from lxml import etree
from docx import Document

# api key
client = openai.OpenAI(api_key="")


# 정규표현식
patterns = {
    "주민등록번호": r"\b\d{6}-\d{7}\b",
    "연락처": r"\b010-\d{4}-\d{4}\b",
    "생년월일": r"\b\d{4}[-/]\d{2}[-/]\d{2}\b",
    "계좌번호": r"\b\d{2,4}-\d{2,4}-\d{2,4}\b",
    "여권번호": r"\b[A-Z]{1}\d{8}\b",
    "이메일": r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b",
    "카드번호": r"\b\d{4}-\d{4}-\d{4}-\d{4}\b", 
}
# 표현식 기반 개인정보 탐지지
def detect_pii_with_regex(content):
    results = {}
    for key, pattern in patterns.items():
        matches = re.findall(pattern, content)
        if matches:
            results[key] = list(set(matches))  # 중복 제거
    return results

# api를 활용한 추가 개인정보 탐지
def detect_sensitive_info_with_chatgpt(content):
    prompt = f"""
    다음 텍스트에서 **문맥을 읽고 실제 개인정보 값만** 찾아주세요.  
    **중요: 불필요한 정보(예: '연락처', '주민등록번호' 같은 분류명)는 반환하지 마세요.**  
    **문맥을 정확히 분석하여 개인정보로 판단되는 값만 반환하세요.**  
    📌 **반드시 JSON 형식으로만 반환하세요.**
    개인정보가 없을 경우 빈 JSON을 반환하세요.

    - **개인정보 예시**  
        - 연락처: "010-1234-5678" ✅  
        - 주민등록번호: "900101-1234567" ✅  
        - 이메일: "test@example.com" ✅  
        - 주소: "서울시 강남구 역삼동" ✅  
        - ❌ 잘못된 예시 (이런 것은 반환하지 마세요)  
            - "연락처" ❌ (값 없이 분류명만)  
            - "신용카드 번호" ❌ (값 없이 분류명만)  

    **반환 형식(JSON):**  
    {{
        "개인정보": {{
            "주소": ["서울시 강남구 역삼동"],
            "연락처": ["010-1234-5678"],
            "이메일": ["example@domain.com"],
            "주민등록번호": ["900101-1234567"],
            "계좌번호": ["1234-5678-9012"]
        }}
    }}
    
    **텍스트 분석 대상:**  
    {content}
    """

    response = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}]
    
    )

    try:
        return json.loads(response.choices[0].message.content)
    except json.JSONDecodeError:
        return {"error": "Invalid JSON from ChatGPT"}

# 마스킹
def apply_masking(content, masking_data):
    for item in masking_data:
        pattern = re.escape(item)  # 특수문자가 포함된 경우...?
        content = re.sub(pattern, "****", content)
    return content

# word 파일에서 텍스트를 추출함
def extract_text_from_word(file_path):
    document = Document(file_path)
    return "\n".join([paragraph.text for paragraph in document.paragraphs])

# xml 파일 처리
def process_xml_file(xml_path, masking_data):
    parser = etree.XMLParser(remove_blank_text=True)
    with open(xml_path, 'rb') as file:
        xml_tree = etree.parse(file, parser)

    for element in xml_tree.xpath("//w:t", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}):
        if element.text:
            element.text = apply_masking(element.text, masking_data)  # xml 파일에서 텍스트 마스킹

    with open(xml_path, 'wb') as file:
        file.write(etree.tostring(xml_tree, pretty_print=True))

# 주석 부분도
def process_comments_xml(comments_path, masking_data):
    if os.path.exists(comments_path):
        process_xml_file(comments_path, masking_data)

# word 파일 마스킹
def mask_sensitive_data_with_images(file_path):
    # 원본 파일에서 개인정보 탐지
    original_content = extract_text_from_word(file_path)
    regex_results = detect_pii_with_regex(original_content)
    chatgpt_results = detect_sensitive_info_with_chatgpt(original_content)

    if "error" in chatgpt_results:
        print("chat gpt 탐지 중 오류 발생:", chatgpt_results["error"])
        return None

    # 개인정보 저장 -> 중복제거
    masking_data = set()
    for key, values in regex_results.items():
        masking_data.update(values)
    for key, values in chatgpt_results.items():
        masking_data.update(values)

    print("탐지된 개인정보:", masking_data)

    # 임시 디렉토리 사용 -> zip 파일 압축 해제
    with TemporaryDirectory() as temp_dir:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # 문서 텍스트 마스킹
        document_xml_path = os.path.join(temp_dir, "word", "document.xml")
        if os.path.exists(document_xml_path):
            process_xml_file(document_xml_path, masking_data)

        # 주석 파일 마스킹 (comments.xml)
        comments_xml_path = os.path.join(temp_dir, "word", "comments.xml")
        process_comments_xml(comments_xml_path, masking_data)

        # 수정된 파일을 다시 .docx로 압축하여 저장
        new_file_path = file_path.replace(".docx", "(masked).docx")
        with zipfile.ZipFile(new_file_path, 'w') as zip_out:
            for foldername, subfolders, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zip_out.write(file_path, arcname)

    return new_file_path


def main():
    input_file = input("📂 마스킹할 Word 파일 경로를 입력하세요: ").strip()
    
    if not os.path.exists(input_file):
        print("❌ 파일이 존재하지 않습니다. 경로를 확인하세요.")
        return

    masked_file = mask_sensitive_data_with_images(input_file)

    if masked_file:
        print(f"✅ 마스킹된 파일이 저장되었습니다: {masked_file}")

if __name__ == "__main__":
    main()
