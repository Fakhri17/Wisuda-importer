import pandas as pd
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from PIL import Image
import re

class GraduationPPTGenerator:
    def __init__(self):
        self.templates = {
            'CUMLAUDE': 'templates/bg_cumlaude.png',
            'SUMMA CUMLAUDE': 'templates/bg_summa.png',
            'Non Predikat': 'templates/bg_non_predikat.png'
        }
        
    def read_excel_data(self, file_path):
        """Read Excel file and return DataFrame"""
        try:
            df = pd.read_excel(file_path)
            print(f"Successfully read {len(df)} rows from {file_path}")
            print(f"Columns: {list(df.columns)}")
            return df
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return None
    
    def get_predikat_template(self, predikat):
        """Determine template based on predikat kelulusan"""
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
        """Find student photo based on NIM in program folder"""
        photo_path = os.path.join('photos', program_folder, f"{nim}_graduation_1.jpg")
        if os.path.exists(photo_path):
            return photo_path
        return None
    
    def extract_seat_position(self, tempat_duduk):
        """Extract seat position for ordering"""
        if pd.isna(tempat_duduk) or tempat_duduk == '':
            return (999, 999, 'Z')  # Put empty seats at the end
        
        # Parse format like "1.1.L" or "1.1.R"
        try:
            parts = str(tempat_duduk).split('.')
            if len(parts) >= 3:
                row = int(parts[0])
                seat = int(parts[1])
                side = parts[2].upper()
                return (row, seat, side)
        except:
            pass
        
        return (999, 999, 'Z')  # Default for invalid format
    
    def create_slide(self, prs, student_data, photo_path):
        """Create a single slide for a student"""
        # Get template based on predikat
        predikat = self.get_predikat_template(student_data.get('PREDIKAT KELULUSAN', ''))
        template_path = self.templates.get(predikat, self.templates['Non Predikat'])
        
        # Add slide with blank layout
        slide_layout = prs.slide_layouts[6]  # Blank layout
        slide = prs.slides.add_slide(slide_layout)
        
        # Set background image
        if os.path.exists(template_path):
            slide.shapes.add_picture(template_path, 0, 0, 
                                   width=prs.slide_width, height=prs.slide_height)
        
        # Add student photo
        if photo_path and os.path.exists(photo_path):
            try:
                # Resize and add photo
                left = Inches(0.5)
                top = Inches(1.5)
                width = Inches(2.5)
                height = Inches(3.5)
                slide.shapes.add_picture(photo_path, left, top, width, height)
            except Exception as e:
                print(f"Error adding photo {photo_path}: {e}")
        
        # Add text boxes for student information
        self.add_student_info(slide, student_data)
        
        return slide
    
    def add_student_info(self, slide, student_data):
        """Add student information to slide"""
        # Faculty (fixed)
        faculty_text = "Telkom University Surabaya"
        
        # Program
        program = student_data.get('PROGRAM STUDI', '')
        
        # Student name
        nama = student_data.get('NAMA MAHASISWA', '')
        
        # NIM
        nim = student_data.get('NIM', '')
        
        # IPK
        ipk = student_data.get('IPK', '')
        
        # TAK
        tak = student_data.get('SKOR TAK', '')
        
        # Dosen Wali
        dosen_wali = student_data.get('Nama Dosen Wali', '')
        
        # Dosen Pembimbing
        dosen_pembimbing1 = student_data.get('Nama Dosen Pembimbing 1', '')
        dosen_pembimbing2 = student_data.get('Nama Dosen Pembimbing 2', '')
        
        # Combine pembimbing names
        if pd.notna(dosen_pembimbing2) and dosen_pembimbing2 != '':
            dosen_pembimbing = f"{dosen_pembimbing1}\n{dosen_pembimbing2}"
        else:
            dosen_pembimbing = dosen_pembimbing1
        
        # Add text boxes
        text_boxes = [
            (faculty_text, Inches(3.5), Inches(1.5), Inches(4), Inches(0.5), 18, True),
            (program, Inches(3.5), Inches(2.0), Inches(4), Inches(0.5), 16, True),
            (nama, Inches(3.5), Inches(2.8), Inches(4), Inches(0.6), 20, True),
            (f"NIM: {nim}", Inches(3.5), Inches(3.5), Inches(4), Inches(0.4), 14, False),
            (f"IPK: {ipk}", Inches(3.5), Inches(3.9), Inches(4), Inches(0.4), 14, False),
            (f"TAK: {tak}", Inches(3.5), Inches(4.3), Inches(4), Inches(0.4), 14, False),
            (f"Dosen Wali: {dosen_wali}", Inches(3.5), Inches(4.7), Inches(4), Inches(0.4), 12, False),
            (f"Dosen Pembimbing:\n{dosen_pembimbing}", Inches(3.5), Inches(5.1), Inches(4), Inches(0.8), 12, False),
        ]
        
        for text, left, top, width, height, font_size, bold in text_boxes:
            if text and str(text).strip() != '' and str(text).strip() != 'nan':
                textbox = slide.shapes.add_textbox(left, top, width, height)
                text_frame = textbox.text_frame
                text_frame.clear()
                p = text_frame.paragraphs[0]
                p.text = str(text)
                p.alignment = PP_ALIGN.LEFT
                
                # Format text
                font = p.font
                font.name = 'Arial'
                font.size = Pt(font_size)
                font.bold = bold
                font.color.rgb = RGBColor(0, 0, 0)  # Black color
    
    def generate_ppt_per_program(self, df, output_dir='output'):
        """Generate separate PPT files for each program"""
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Group by program
        programs = df['PROGRAM STUDI'].unique()
        
        for program in programs:
            if pd.isna(program) or program == '':
                continue
                
            print(f"\nProcessing program: {program}")
            
            # Filter data for this program
            program_data = df[df['PROGRAM STUDI'] == program].copy()
            
            # Sort by seat position
            program_data['seat_sort'] = program_data['TEMPAT DUDUK'].apply(self.extract_seat_position)
            program_data = program_data.sort_values('seat_sort')
            
            # Create presentation
            prs = Presentation()
            
            # Add slides for each student
            for _, student in program_data.iterrows():
                nim = student.get('NIM', '')
                photo_path = self.find_student_photo(nim, program)
                
                if photo_path:
                    print(f"  Adding slide for {student.get('NAMA MAHASISWA', '')} (NIM: {nim})")
                else:
                    print(f"  Warning: Photo not found for {student.get('NAMA MAHASISWA', '')} (NIM: {nim})")
                
                self.create_slide(prs, student, photo_path)
            
            # Save presentation
            safe_program_name = re.sub(r'[^\w\s-]', '', str(program)).strip()
            safe_program_name = re.sub(r'[-\s]+', '_', safe_program_name)
            output_file = os.path.join(output_dir, f"{safe_program_name}.pptx")
            
            prs.save(output_file)
            print(f"  Saved: {output_file} ({len(program_data)} slides)")
    
    def process_graduation_data(self, excel_file, output_dir='output'):
        """Main function to process graduation data"""
        print(f"Processing graduation data from: {excel_file}")
        
        # Read Excel data
        df = self.read_excel_data(excel_file)
        if df is None:
            return
        
        # Check required columns
        required_columns = [
            'PROGRAM STUDI', 'NAMA MAHASISWA', 'NIM', 'IPK', 'SKOR TAK',
            'Nama Dosen Wali', 'Nama Dosen Pembimbing 1', 'Nama Dosen Pembimbing 2',
            'PREDIKAT KELULUSAN', 'TEMPAT DUDUK'
        ]
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"Warning: Missing columns: {missing_columns}")
            print(f"Available columns: {list(df.columns)}")
        
        # Show predikat distribution
        if 'PREDIKAT KELULUSAN' in df.columns:
            print(f"\nPredikat distribution:")
            predikat_counts = df['PREDIKAT KELULUSAN'].value_counts()
            for predikat, count in predikat_counts.items():
                template = self.get_predikat_template(predikat)
                print(f"  {predikat}: {count} students -> {template} template")
        
        # Generate PPT files per program
        self.generate_ppt_per_program(df, output_dir)
        
        print(f"\nProcessing completed! Check the '{output_dir}' folder for generated PPT files.")

def main():
    """Main function"""
    generator = GraduationPPTGenerator()
    
    # Process both Excel files with organized folder structure
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
            
            # Create organized output directory under root 'output/'
            output_dir = os.path.join('output', folder_name)
            generator.process_graduation_data(excel_file, output_dir)
        else:
            print(f"File not found: {excel_file}")

if __name__ == "__main__":
    main()
