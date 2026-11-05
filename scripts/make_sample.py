"""Create a small valid .pptx for the CI smoke test."""
import sys
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "ooxml-validate smoke test"
slide.placeholders[1].text = "If this passes, the validator runs."
slide.shapes.add_textbox(Inches(1), Inches(5), Inches(4), Inches(1)).text_frame.text = "Hello"
prs.save(sys.argv[1] if len(sys.argv) > 1 else "sample.pptx")
