import os
import re
import pandas as pd
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

class GraduationPPTGenerator:
    DPI = 96  # konsisten dengan PowerPoint

    def __init__(self):
        self.templates = {
            'CUMLAUDE': 'templates/bg_cumlaude.png',
            'SUMMA CUMLAUDE': 'templates/bg_summa.png',
            'Non Predikat': 'templates/bg_non_predikat.png'
        }

    # =========================
    # Helpers ukuran & gambar
    # =========================
    def _set_slide_size_to_image_exact(self, prs, image_path, dpi=None):
        """Sesuaikan ukuran slide PERSIS dengan ukuran gambar (pixel -> inch @DPI)."""
        if dpi is None:
            dpi = self.DPI
        if not image_path or not os.path.exists(image_path):
            return
        with Image.open(image_path) as img:
            w_px, h_px = img.size
        prs.slide_width  = Inches(w_px / dpi)
        prs.slide_height = Inches(h_px / dpi)

    def _set_background_image(self, slide, image_path):
        """Pasang gambar sebagai latar dengan menambahkan picture di (0,0) ukuran native."""
        if not image_path or not os.path.exists(image_path):
            return
        try:
            # Native size; asalkan ukuran slide sudah diset 1:1 dengan gambar, ini akan full-bleed
            slide.shapes.add_picture(image_path, 0, 0)
        except Exception as e:
            print(f"Error setting background image {image_path}: {e}")

    def _add_picture_fit(self, slide, image_path, left, top, frame_width, frame_height):
        """Tambahkan gambar agar pas di dalam frame tanpa distorsi (centered)."""
        try:
            with Image.open(image_path) as img:
                img_w, img_h = img.size
        except Exception as e:
            print(f"Error opening image {image_path}: {e}")
            return None

        frame_ratio = float(frame_width) / float(frame_height) if frame_height else 1.0
        img_ratio = img_w / img_h if img_h else 1.0

        if img_ratio > frame_ratio:
            # Fit to width
            width = frame_width
            height = int(float(width) / img_ratio)
        else:
            # Fit to height
            height = frame_height
            width = int(float(height) * img_ratio)

        offset_left = left + (frame_width - width) / 2
        offset_top = top + (frame_height - height) / 2

        try:
            return slide.shapes.add_picture(image_path, offset_left, offset_top, width=width, height=height)
        except Exception as e:
            print(f"Error adding fitted picture {image_path}: {e}")
            return None

    def _add_picture_center_no_resize(self, slide, image_path, center_x, center_y, dpi=None):
        """
        Tambahkan gambar pada ukuran aslinya (@DPI) dan posisikan TEPAT di tengah (tanpa resize).
        center_x / center_y dalam EMU (pakai Inches(...) saat panggil).
        """
        if dpi is None:
            dpi = self.DPI
        if not image_path or not os.path.exists(image_path):
            return None

        with Image.open(image_path) as img:
            w_px, h_px = img.size

        width = Inches(w_px / dpi)
        height = Inches(h_px / dpi)
        left = center_x - width / 2
        top = center_y - height / 2

        try:
            return slide.shapes.add_picture(image_path, left, top, width=width, height=height)
        except Exception as e:
            print(f"Error adding centered picture {image_path}: {e}")
            return None

    # =========================
    # Data helpers
    # =========================
    def read_excel_data(self, file_path):
        """Read Excel file and return DataFrame."""
        try:
            df = pd.read_excel(file_path)
            print(f"Successfully read {len(df)} rows from {file_path}")
            print(f"Columns: {list(df.columns)}")
            return df
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return None

    def get_predikat_template(self, predikat):
        """Determine template based on predikat kelulusan."""
        if pd.isna(predikat) or predikat == '':
            return 'Non Predikat'
        predikat_str = str(predikat).lower().strip()
        if 'summa' in predikat_str and 'cumlaude' in predikat_str:
            return 'SUMMA CUMLAUDE'
        elif 'cumlaude' in predikat_str and 'summa' not in predikat_str:
            return 'CUMLAUDE'
        else:
            return 'Non Predikat'

    def find_student_photo(self, nim, program_folder):
        """Find student photo based on NIM in program folder."""
        photo_path = os.path.join('photos', program_folder, f"{nim}_graduation_1.jpg")
        return photo_path if os.path.exists(photo_path) else None

    def extract_seat_position(self, tempat_duduk):
        """Extract seat position for ordering (format '1.1.L')."""
        if pd.isna(tempat_duduk) or str(tempat_duduk).strip() == '':
            return (999, 999, 'Z')
        try:
            parts = str(tempat_duduk).split('.')
            if len(parts) >= 3:
                row = int(parts[0]); seat = int(parts[1]); side = parts[2].upper()
                return (row, seat, side)
        except Exception:
            pass
        return (999, 999, 'Z')

    # =========================
    # Slide builders
    # =========================
    def create_slide(self, prs, student_data, photo_path):
        """Create a single slide for a student."""
        predikat = self.get_predikat_template(student_data.get('PREDIKAT KELULUSAN', ''))
        template_path = self.templates.get(predikat, self.templates['Non Predikat'])

        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Background full-bleed TANPA distorsi
        if os.path.exists(template_path):
            self._set_background_image(slide, template_path)

        # Posisi frame foto (area tempat foto seharusnya berada)
        frame_left, frame_top = Inches(0.5), Inches(1.5)
        frame_w, frame_h = Inches(2.5), Inches(3.5)
        center_x = frame_left + frame_w / 2
        center_y = frame_top + frame_h / 2

        # Foto: TARUH DI TENGAH TANPA RESIZE; kalau kebesaran, fallback ke fit.
        if photo_path and os.path.exists(photo_path):
            try:
                pic = self._add_picture_center_no_resize(slide, photo_path, center_x, center_y)
                if pic is not None:
                    # Jika melebihi frame, hapus dan fit ke frame
                    out_left = (pic.left < frame_left)
                    out_top = (pic.top < frame_top)
                    out_right = (pic.left + pic.width > frame_left + frame_w)
                    out_bottom = (pic.top + pic.height > frame_top + frame_h)
                    if out_left or out_top or out_right or out_bottom:
                        pic._element.getparent().remove(pic._element)
                        self._add_picture_fit(slide, photo_path, frame_left, frame_top, frame_w, frame_h)
            except Exception as e:
                print(f"Error adding photo {photo_path}: {e}")

        # Teks info mahasiswa & dosen
        self.add_student_info(slide, student_data)

        return slide

    def _add_textbox(self, slide, text, left, top, width, height, font_size=18, bold=False, upper=True):
        """Utility untuk menambah textbox konsisten."""
        if not text or str(text).strip().lower() == 'nan':
            return
        content = str(text).upper() if upper else str(text)
        tb = slide.shapes.add_textbox(left, top, width, height)
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = content
        p.alignment = PP_ALIGN.LEFT

        # pastikan ada run
        if len(p.runs) == 0:
            run = p.add_run()
            run.text = content
            p.text = ''

        for run in p.runs:
            run.font.name = 'Arial'
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.color.rgb = RGBColor(0, 0, 0)

    def add_student_info(self, slide, student_data):
        """Add student information to slide."""
        program = student_data.get('PROGRAM STUDI', '')
        nama = student_data.get('NAMA MAHASISWA', '')
        nim = student_data.get('NIM', '')
        ipk = student_data.get('IPK', '')
        tak = student_data.get('SKOR TAK', '')
        dosen_wali = student_data.get('Nama Dosen Wali', '')

        dosen_pembimbing1 = student_data.get('Nama Dosen Pembimbing 1', '')
        dosen_pembimbing2 = student_data.get('Nama Dosen Pembimbing 2', '')
        pembimbing_names = []
        if pd.notna(dosen_pembimbing1) and str(dosen_pembimbing1).strip() != '':
            pembimbing_names.append(str(dosen_pembimbing1))
        if pd.notna(dosen_pembimbing2) and str(dosen_pembimbing2).strip() != '':
            pembimbing_names.append(str(dosen_pembimbing2))

        # Informasi utama
        self._add_textbox(slide, program, Inches(3.5), Inches(2.0), Inches(4), Inches(0.5), font_size=14, bold=True)
        self._add_textbox(slide, nama, Inches(3.5), Inches(2.8), Inches(4), Inches(0.6), font_size=19, bold=True)
        self._add_textbox(slide, f"NIM : {nim}", Inches(3.5), Inches(3.5), Inches(4), Inches(0.4), font_size=16, bold=True)
        self._add_textbox(slide, f"IPK : {ipk} – TAK : {tak}", Inches(3.5), Inches(3.9), Inches(4), Inches(0.4), font_size=16, bold=True)
        self._add_textbox(slide, f"DOSEN WALI : {dosen_wali}", Inches(3.5), Inches(4.7), Inches(4), Inches(0.4), font_size=14, bold=True)

        # Label "DOSEN PEMBIMBING :" dan nama-nama dipisah shape
        label_left = Inches(3.5)
        label_top = Inches(5.1)
        label_width = Inches(2.2)
        label_height = Inches(0.4)

        # Label
        label_box = slide.shapes.add_textbox(label_left, label_top, label_width, label_height)
        label_tf = label_box.text_frame
        label_tf.clear()
        p_label = label_tf.paragraphs[0]
        p_label.alignment = PP_ALIGN.LEFT
        run_label = p_label.add_run()
        run_label.text = "DOSEN PEMBIMBING :"
        run_label.font.name = 'Arial'
        run_label.font.size = Pt(14)
        run_label.font.bold = True
        run_label.font.color.rgb = RGBColor(0, 0, 0)

        # Names box – dimulai tepat setelah titik dua (pakai shape terpisah)
        names_left = label_left + label_width
        names_top = label_top
        names_width = Inches(4)
        names_height = Inches(0.8)

        names_box = slide.shapes.add_textbox(names_left, names_top, names_width, names_height)
        names_tf = names_box.text_frame
        names_tf.clear()

        if len(pembimbing_names) > 0:
            # baris pertama
            p_name = names_tf.paragraphs[0]
            p_name.alignment = PP_ALIGN.LEFT
            run_n1 = p_name.add_run()
            run_n1.text = str(pembimbing_names[0]).upper()
            run_n1.font.name = 'Arial'
            run_n1.font.size = Pt(14)
            run_n1.font.bold = True
            run_n1.font.color.rgb = RGBColor(0, 0, 0)

            # baris berikutnya
            for extra in pembimbing_names[1:]:
                p_more = names_tf.add_paragraph()
                p_more.alignment = PP_ALIGN.LEFT
                run_more = p_more.add_run()
                run_more.text = str(extra).upper()
                run_more.font.name = 'Arial'
                run_more.font.size = Pt(14)
                run_more.font.bold = True
                run_more.font.color.rgb = RGBColor(0, 0, 0)

    # =========================
    # Pipeline
    # =========================
    def generate_ppt_per_program(self, df, output_dir='output'):
        """Generate separate PPT files for each program."""
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        programs = df['PROGRAM STUDI'].unique()

        for program in programs:
            if pd.isna(program) or str(program).strip() == '':
                continue

            print(f"\nProcessing program: {program}")
            program_data = df[df['PROGRAM STUDI'] == program].copy()

            # Sort by seat position
            program_data['seat_sort'] = program_data['TEMPAT DUDUK'].apply(self.extract_seat_position)
            program_data = program_data.sort_values('seat_sort')

            prs = Presentation()

            # Tentukan template agar ukuran slide match EXACT background
            try:
                first_row = program_data.iloc[0]
                first_template_name = self.get_predikat_template(first_row.get('PREDIKAT KELULUSAN', ''))
                first_template_path = self.templates.get(first_template_name, self.templates['Non Predikat'])
            except Exception:
                first_template_path = self.templates['Non Predikat']

            self._set_slide_size_to_image_exact(prs, first_template_path)

            # Tambahkan slide per mahasiswa
            for _, student in program_data.iterrows():
                nim = student.get('NIM', '')
                photo_path = self.find_student_photo(nim, program)
                if photo_path:
                    print(f"  Adding slide for {student.get('NAMA MAHASISWA', '')} (NIM: {nim})")
                else:
                    print(f"  Warning: Photo not found for {student.get('NAMA MAHASISWA', '')} (NIM: {nim})")
                self.create_slide(prs, student, photo_path)

            # Simpan file
            safe_program_name = re.sub(r'[^\w\s-]', '', str(program)).strip()
            safe_program_name = re.sub(r'[-\s]+', '_', safe_program_name)
            output_file = os.path.join(output_dir, f"{safe_program_name}.pptx")
            prs.save(output_file)
            print(f"  Saved: {output_file} ({len(program_data)} slides)")

    def process_graduation_data(self, excel_file, output_dir='output'):
        """Main function to process graduation data."""
        print(f"Processing graduation data from: {excel_file}")
        df = self.read_excel_data(excel_file)
        if df is None:
            return

        required_columns = [
            'PROGRAM STUDI', 'NAMA MAHASISWA', 'NIM', 'IPK', 'SKOR TAK',
            'Nama Dosen Wali', 'Nama Dosen Pembimbing 1', 'Nama Dosen Pembimbing 2',
            'PREDIKAT KELULUSAN', 'TEMPAT DUDUK'
        ]
        missing = [c for c in required_columns if c not in df.columns]
        if missing:
            print(f"Warning: Missing columns: {missing}")
            print(f"Available columns: {list(df.columns)}")

        if 'PREDIKAT KELULUSAN' in df.columns:
            print("\nPredikat distribution:")
            counts = df['PREDIKAT KELULUSAN'].value_counts()
            for p, c in counts.items():
                t = self.get_predikat_template(p)
                print(f"  {p}: {c} students -> {t} template")

        self.generate_ppt_per_program(df, output_dir)
        print(f"\nProcessing completed! Check the '{output_dir}' folder for generated PPT files.")

def main():
    generator = GraduationPPTGenerator()

    # Daftar file excel & subfolder output
    excel_files = [
        ('wisuda_pagi.xlsx', 'Wisuda Pagi'),
        ('wisuda_siang.xlsx', 'Wisuda Siang')
    ]

    for excel_file, folder_name in excel_files:
        if os.path.exists(excel_file):
            print(f"\n{'='*50}")
            print(f"Processing: {excel_file}")
            print(f"Output folder: {folder_name}")
            print(f"{'='*50}")
            output_dir = os.path.join('output', folder_name)
            generator.process_graduation_data(excel_file, output_dir)
        else:
            print(f"File not found: {excel_file}")

if __name__ == "__main__":
    main()
