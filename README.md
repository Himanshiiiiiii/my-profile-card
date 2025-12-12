# my-profile-card
This is my sample profile card
-----------------
#!/usr/bin/env python3
"""
pptx_extract_md.py

Usage:
    python pptx_extract_md.py test1.pptx output_folder
"""

import sys
import os
import shutil
import zipfile
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def ensure_outdir(path):
    os.makedirs(path, exist_ok=True)


def save_file_bytes(outdir, name, data_bytes):
    path = os.path.join(outdir, name)
    with open(path, "wb") as f:
        f.write(data_bytes)
    return path


def extract_text_and_tables(prs):
    slides_output = []

    for i, slide in enumerate(prs.slides, start=1):
        slide_items = []

        # collect all shapes with coordinates
        for shape in slide.shapes:
            top = getattr(shape, "top", 0)
            left = getattr(shape, "left", 0)
            slide_items.append((top, left, shape))

        # sort shapes by top → left (visual reading order)
        slide_items.sort(key=lambda x: (x[0], x[1]))

        slide_text = []
        slide_tables = []

        for _, _, shape in slide_items:

            # extract text
            if shape.has_text_frame:
                text = shape.text.strip()
                if text:
                    slide_text.append(text)

            # extract table
            if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                tbl = shape.table
                table_data = []
                for r in range(len(tbl.rows)):
                    row = []
                    for c in range(len(tbl.columns)):
                        row.append(tbl.cell(r, c).text.strip())
                    table_data.append(row)
                slide_tables.append(table_data)

        slides_output.append({
            "slide_index": i,
            "texts": slide_text,
            "tables": slide_tables
        })

    return slides_output


def extract_images(prs, outdir):
    images = []
    counter = 1
    for i, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                img = shape.image
                ext = img.ext
                name = f"slide_{i}_image_{counter}.{ext}"
                save_file_bytes(outdir, name, img.blob)
                images.append((i, name))
                counter += 1
    return images


def extract_charts(prs, outdir):
    charts = []
    chart_counter = 1

    for i, slide in enumerate(prs.slides, start=1):

        for shape in slide.shapes:
            try:
                chart = shape.chart
            except Exception:
                continue

            series_list = []
            categories = None

            try:
                if hasattr(chart, "category_axis") and chart.category_axis.categories:
                    categories = [c.label for c in chart.category_axis.categories]
            except Exception:
                pass

            try:
                if categories is None and chart.plots and chart.plots[0].categories:
                    categories = [c.label for c in chart.plots[0].categories]
            except Exception:
                pass

            try:
                for s in chart.series:
                    name = s.name
                    vals = list(s.values)
                    series_list.append({"name": name, "values": vals})
            except Exception as e:
                series_list.append({"error": str(e)})

            charts.append({
                "slide_index": i,
                "chart_index": chart_counter,
                "categories": categories,
                "series": series_list
            })

            chart_counter += 1

    return charts


def extract_embedded_files(pptx_path, outdir):
    emb_files = []
    with zipfile.ZipFile(pptx_path, 'r') as z:
        for zi in z.infolist():
            if zi.filename.startswith("ppt/embeddings/"):
                name = os.path.basename(zi.filename)
                out = os.path.join(outdir, name)
                with z.open(zi) as src, open(out, "wb") as dst:
                    shutil.copyfileobj(src, dst)
                emb_files.append(name)
    return emb_files


def write_markdown(outdir, pptx_name, slides, images, charts, embedded):
    md_path = os.path.join(outdir, "extracted_output.md")

    with open(md_path, "w", encoding="utf-8") as md:

        md.write(f"# Extracted Content from `{pptx_name}`\n\n")
        md.write(f"Total Slides: **{len(slides)}**\n\n---\n")

        for slide in slides:
            md.write(f"## Slide {slide['slide_index']}\n")

            if slide["texts"]:
                md.write("### Text Content\n")
                for t in slide["texts"]:
                    md.write(f"- {t}\n")
                md.write("\n")

            for t_idx, table in enumerate(slide["tables"], start=1):
                md.write(f"### Table {t_idx}\n")
                md.write("| " + " | ".join(table[0]) + " |\n")
                md.write("|" + " --- |" * len(table[0]) + "\n")
                for row in table[1:]:
                    md.write("| " + " | ".join(row) + " |\n")
                md.write("\n")

            slide_imgs = [img for s, img in images if s == slide["slide_index"]]
            if slide_imgs:
                md.write("### Images\n")
                for img in slide_imgs:
                    md.write(f"- {img}\n")
                md.write("\n")

            slide_charts = [c for c in charts if c["slide_index"] == slide["slide_index"]]
            for c in slide_charts:
                md.write(f"### Chart {c['chart_index']}\n")
                md.write(f"- **Categories:** {c['categories']}\n")
                md.write("- **Series:**\n")
                for s in c["series"]:
                    md.write(f"  - {s}\n")
                md.write("\n")

            md.write("---\n")

        if embedded:
            md.write("## Embedded Files\n")
            for f in embedded:
                md.write(f"- {f}\n")

    return md_path


def main(pptx_path, outdir):
    ensure_outdir(outdir)
    prs = Presentation(pptx_path)

    slides = extract_text_and_tables(prs)
    images = extract_images(prs, outdir)
    charts = extract_charts(prs, outdir)
    embedded = extract_embedded_files(pptx_path, outdir)

    md_path = write_markdown(outdir, os.path.basename(pptx_path), slides, images, charts, embedded)

    print("Markdown file created:", md_path)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python pptx_extract_md.py test1.pptx output_folder")
        sys.exit(1)

    main(sys.argv[1], sys.argv[2])
