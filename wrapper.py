import os
from docx import Document

TITLE = "# Temas de Historia de la Computación"
OUTPUT_MD = "temas.md"
DIRECTORY = "./Lista de Temas de HC 2025"

SECTIONS_MAP = {
    0: "Explicación del tema",
    1: "Cantidad de integrantes",
    2: "Cantidad de equipos que pueden desarrollar el tema",
}


def create_markdown_file(filename):
    with open(filename, "w", encoding="utf-8") as file:
        file.write(f"{TITLE}\n\n")
    return filename


def write_to_markdown(filename, content):
    with open(filename, "a", encoding="utf-8") as file:
        file.write(content)


def process_directories(base_dir, output_file):
    for folder_name in os.listdir(base_dir):
        folder_path = os.path.join(base_dir, folder_name)
        if os.path.isdir(folder_path):
            try:
                process_directory(folder_path, folder_name, output_file)
            except Exception as e:
                print(f"Error al procesar la carpeta {folder_name}: {e}")


def process_directory(dir_path, dir_name, output_file):
    docx_file = find_docx_file(dir_path)
    if not docx_file:
        print(f"No se encontró un archivo .docx en {dir_path}")
        return
    content = extract_docx_content(docx_file)

    if not content:
        print(f"No se pudo extraer contenido del archivo {docx_file}")
        return

    write_to_markdown(output_file, f"## {dir_name}\n\n")
    write_sections(output_file, content)


def find_docx_file(dir_path):
    for file_name in os.listdir(dir_path):
        if file_name.endswith(".docx"):
            return os.path.join(dir_path, file_name)
    return None


def extract_docx_content(docx_path):
    try:
        doc = Document(docx_path)
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        return paragraphs
    except Exception as e:
        print(f"Error al leer el archivo .docx {docx_path}: {e}")
        return None


def write_sections(output_file, content):
    for key, paragraph in enumerate(content):
        section_title = SECTIONS_MAP.get(key)
        if section_title:
            write_to_markdown(output_file, f"### {section_title}\n{paragraph}\n\n")
    return True


def main():
    output_file = create_markdown_file(OUTPUT_MD)

    process_directories(DIRECTORY, output_file)

    print(f"Información guardada en {OUTPUT_MD}")


if __name__ == "__main__":
    main()
