# DocxComparator

The **DocxComparator** is a Python program that enables you to compare, align, and improve text from two .docx files. It uses spaCy NLP to generate an updated .docx file with improved text based on the alignment results.

## Features

- Compare and compute the similarity score between two .docx files.
- Align the text from both files to identify similar and dissimilar sentences.
- Analyze and improve the text alignment to suggest enhanced sentences.
- Generate a new .docx file containing the improved text.

## Requirements

- Python 3.x
- `docx` library
- `spacy` library with English language model (en_core_web_sm)

## Installation

1. Clone the repository or download the zip file.
2. Install the required libraries by running `pip install docx spacy` and `python -m spacy download en_core_web_sm`.

## Usage

1. Create an instance of the `DocxComparator` class with the file paths of the two .docx files to be compared.
2. Run the `run()` method to perform the comparison, alignment, and text improvement.
3. The program will print the similarity score, aligned text, and updated text, and it will save the improved text in a new .docx file.

Example:

```python
from DocxComparator import DocxComparator

file1_path = 'path/to/file1.docx'
file2_path = 'path/to/file2.docx'

comparator = DocxComparator(file1_path, file2_path)
comparator.run()
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [spaCy](https://spacy.io/) - An open-source natural language processing library.
- [python-docx](https://python-docx.readthedocs.io/) - A library to work with .docx files in Python.
- [Inspiration](https://github.com/suyash248/document-text-comparison) - Inspiration for this project came from a similar document text comparison repository.
