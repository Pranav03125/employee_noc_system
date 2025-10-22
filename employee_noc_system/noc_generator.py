import os
from docx import Document

class EmployeeNOCGenerator:
    """
    Generates personalized NOC documents by inserting user input
    next to 'Full Name:', 'Job Title:', and 'Department:' exactly as in the template.
    """

    def __init__(self, template_path: str):
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template not found: {template_path}")
        self.template_path = template_path

    def _replace_field(self, text: str, key: str, value: str) -> str:
        """
        Replaces or appends user input after 'key:' while keeping the key visible.
        Example:
            'Full Name:' → 'Full Name: Arjun Kumar'
        """
        if key.lower() in text.lower():
            # Split only once after the first colon
            parts = text.split(":", 1)
            if len(parts) == 2:
                return f"{parts[0]}: {value}"
            else:
                return f"{text.strip()}: {value}"
        return text

    def generate_noc(self, full_name: str, job_title: str, department: str, output_dir: str = "generated_noc") -> str:
        """
        Generates and saves a personalized NOC file.
        """
        doc = Document(self.template_path)

        replacements = {
            "Full Name": full_name.strip(),
            "Job Title": job_title.strip(),
            "Department": department.strip()
        }

        # Process all paragraphs in document
        for para in doc.paragraphs:
            original_text = para.text
            for key, value in replacements.items():
                if key.lower() in original_text.lower():
                    para.text = self._replace_field(original_text, key, value)

        # Process text inside tables (some docs store fields inside tables)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        original_text = para.text
                        for key, value in replacements.items():
                            if key.lower() in original_text.lower():
                                para.text = self._replace_field(original_text, key, value)

        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)
        safe_name = full_name.replace(" ", "_")
        output_path = os.path.join(output_dir, f"NOC_{safe_name}.docx")

        doc.save(output_path)
        print(f"NOC generated successfully for {full_name}: {output_path}")
        return output_path


if __name__ == "__main__":
    generator = EmployeeNOCGenerator("NDA-1.docx")

    full_name = input("Enter Full Name: ").strip()
    job_title = input("Enter Job Title: ").strip()
    department = input("Enter Department: ").strip()

    generator.generate_noc(full_name, job_title, department)
