import unittest
import os
from book_writer import BookWriter
from docx import Document

class TestBookWriter(unittest.TestCase):
    def test_book_creation_and_format(self):
        theme = "Test Book KDP"
        filename = "Test_Book_KDP.docx"
        
        if os.path.exists(filename):
            os.remove(filename)
            
        writer = BookWriter(theme)
        writer.create_book()
        writer.save_document() # Need to call this explicitly now
        
        self.assertTrue(os.path.exists(filename))
        
        # Load and verify format
        doc = Document(filename)
        section = doc.sections[0]
        
        # Verify 6x9 inches
        self.assertAlmostEqual(section.page_width.inches, 6.0, places=1)
        self.assertAlmostEqual(section.page_height.inches, 9.0, places=1)
        
        # Check for Sumário and Title
        headings = [p.text for p in doc.paragraphs if p.style.name.startswith('Heading') or p.style.name == 'Title']
        self.assertIn("Sumário", headings)
        self.assertIn(theme, headings)
        
        # Cleanup
        os.remove(filename)

if __name__ == "__main__":
    unittest.main()
