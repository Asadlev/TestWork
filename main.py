from docxtpl import DocxTemplate
from docx.shared import Inches


class DocumentBuilder:
    def __init__(self, templates_order):
        self.templates_order = templates_order
        self.templates = {}
        self.dataset = {}

    def load_templates(self):
        for template_file, _ in self.templates_order.items():
            template = DocxTemplate(template_file)
            self.templates[template_file] = template

    def load_dataset(self, dataset):
        self.dataset = dataset

    def validate_templates(self):
        for template_file, _ in self.templates_order.items():
            template = self.templates[template_file]
            for key, value in self.dataset.items():
                if not template.contains(key):
                    print(f"Template {template_file} does not contain data for key '{key}'")
            for variable in template.get_missing_variables(self.dataset):
                print(f"Variable '{variable}' is missing in dataset")

    def build_document(self):
        document = DocxTemplate("output.docx")
        for template_file, order in sorted(self.templates_order.items(), key=lambda x: x[1]):
            template = self.templates[template_file]
            document.attach(template)
            document.merge(**self.dataset)
        document.save("output.docx")


if __name__ == "__main__":
    templates_order = {
        "heading_1.docx": 0,
        "heading_2.docx": 1,
        "body_1.docx": 2,
        "body_2.docx": 3,
        "bottom_1.docx": 4,
        "bottom_2.docx": 5
    }

    dataset = {
        "variable1": "value1",
        "variable2": "value2",
        # Добавьте остальные переменные из датасета
    }

    builder = DocumentBuilder(templates_order)
    builder.load_templates()
    builder.load_dataset(dataset)
    builder.validate_templates()
    builder.build_document()