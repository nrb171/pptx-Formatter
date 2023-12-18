from pptx.oxml.xmlchemy import OxmlElement
import json
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

fileToFormat = "/Users/nrb171/Library/CloudStorage/OneDrive-ThePennsylvaniaStateUniversity/PSU/Instructing/SP24 - METEO273/lectures/SP24/02 startProgramming.pptx"


def apply_format(run, format_spec):
    try:
        run.font.name = format_spec['font_name']

        # run.font.size = run.font.size
        # run.font.bold = format_spec['font_bold']
        # run.font.italic = format_spec['font_italic']

        run.font.color.rgb = RGBColor(*format_spec['font_color'])
    except TypeError:
        return


def format_run(paragraph, word, format_spec):
    new_runs = []  # List to store new runs

    for run in paragraph.runs:
        split_text = run.text.split(word)

        try:
            if split_text[0][-1] == ' ' or split_text[1][0] == ' ':
                partOfWord = False
            else:
                partOfWord = True
        except:
            partOfWord = False

        if word in run.text and partOfWord == False:

            # Iterate through each part in split_text
            for i, part in enumerate(split_text):
                # Add the non-word part
                if part:
                    new_run = paragraph.add_run()
                    new_run.text = part
                    new_run = _copy_run_formatting(new_run, run)
                    new_runs.append(new_run)

                # Add the formatted word, except for the last part if it's empty
                if i < len(split_text) - 1:
                    formatted_run = paragraph.add_run()
                    formatted_run.text = word
                    apply_format(formatted_run, format_spec)
                    new_runs.append(formatted_run)
        elif run.text != '':
            # If the word is not in the run, just copy the run as is
            new_run = paragraph.add_run()
            new_run.text = run.text
            new_run = _copy_run_formatting(new_run, run)
            new_runs.append(new_run)

    # Clear original runs and update paragraph with new runs
    paragraph.clear()
    for new_run in new_runs:
        # Add the new run to the paragraph's xml
        run = paragraph.add_run()
        run.text = new_run.text
        run = _copy_run_formatting(run, new_run)

    # Clear the temporary new_runs list

    return paragraph


def _copy_run_formatting(new_run, original_run):
    # Copy formatting from original run to new run
    try:
        new_run.font.name = original_run.font.name
        new_run.font.size = original_run.font.size
        new_run.font.bold = original_run.font.bold
        new_run.font.italic = original_run.font.italic
        new_run.font.color.rgb = original_run.font.color.rgb
        return new_run
    except:
        return new_run


# Load the presentation and formatting rules


# Iterate through each slide and textbox
# Iterate through each slide and textbox
with open('words.json', 'r') as file:
    formatting_rules = json.load(file)

pptx = Presentation(fileToFormat)

for slide in pptx.slides:
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for rule in formatting_rules['words']:
                    # Iterate over each word in the "words" list
                    for word in rule['words']:
                        if word in paragraph.text.lower():

                            paragraph = format_run(
                                paragraph, word, rule['format'])
                            print(paragraph.text)

        # Save the formatted presentation
pptx.save(fileToFormat.replace('.pptx', '_formatted.pptx'))
