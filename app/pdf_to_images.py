from pathlib import Path
import fitz  # PyMuPDF


def render_pdf_pages_to_images(pdf_path: str, output_dir: str = "data/pdf_pages", dpi: int = 200) -> list[str]:
    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)
    image_paths = []

    zoom = dpi / 72
    matrix = fitz.Matrix(zoom, zoom)

    for page_index in range(len(doc)):
        page = doc[page_index]
        pix = page.get_pixmap(matrix=matrix, alpha=False)

        image_path = output / f"page_{page_index + 1:03}.png"
        pix.save(str(image_path))
        image_paths.append(str(image_path))

    doc.close()
    return image_paths