
# **INTRODUCTION TO OFFICER PACKAGE**

The `officer` package in R is a powerful tool for creating and manipulating Microsoft Word and PowerPoint documents programmatically. It provides a way to generate reports, presentations, and other documents with dynamic content, formatting, and styling. The package enables you to automate the process of document creation, making it easier to generate reproducible reports, documents, and presentations directly from your R code. Here's a brief introduction to the key features and concepts of the officer package:

**Key Features**:

- Create and Modify Documents: 
With the officer package, you can create new Word documents and PowerPoint presentations from scratch or modify existing ones. This makes it suitable for both report generation and presentation creation.

- Dynamic Content: 
The package allows you to incorporate dynamic content, such as text, tables, graphs, images, and other elements, into your documents. This is particularly useful when you need to create reports based on changing data.

- Formatting and Styling: 
Officer provides extensive options for formatting and styling your documents. You can control font styles, sizes, colors, alignments, indentation, and more, to achieve the desired visual appearance.

- Templates: 
You can define custom styles and templates that maintain consistent formatting across different sections of your documents. This ensures a professional and polished look.

- Tables and Graphs: 
Creating tables and inserting graphs is straightforward using the officer package. You can use data from R to populate tables and generate graphs using packages like ggplot2.

- Reusable Components: 
You can create reusable components such as headers, footers, cover pages, and table of contents to maintain consistency across different documents.

- Export Formats: 
The package allows you to export your final documents in various formats, including Word documents (.docx) and PowerPoint presentations (.pptx). You can also convert Word documents to PDF using external tools.

**Basic Concepts**:

- Document Object: 
The central concept in the `officer` package is the document object (e.g., docx_document or pptx_document), which serves as a canvas for building your document or presentation.

- Cursor: 
A cursor is used to navigate through the document, allowing you to add content at specific locations. It's similar to a pointer that keeps track of where new content should be added.

- Paragraph Styles: 
Paragraph styles define the appearance of text paragraphs, including font, size, color, alignment, and indentation. You can create and apply custom paragraph styles to maintain consistent formatting.

- Formatting Paragraphs: 
You can format paragraphs using `body_add_par()` function. You can also define paragraphs with specific styles, as well as use special characters for line breaks and page breaks.

- Content Addition: 
The officer package offers functions like `body_add_par()`, `body_add_table()`, and `body_add_plot()` to add different types of content to the document.

- Styles and Templates: 
You can define and apply styles for different document elements, such as paragraphs, headers, and footers. This helps maintain a consistent look throughout the document.

Learn more about [Officer](https://davidgohel.github.io/officer/)

**Getting Started**:

### Install and load the library

```r
install.packages("officer")
library(officer)
library(dplyr)
```

```r
word_doc <- read_docx() #no file specified in () so, it's blank
word_doc <- word_doc %>% 
body_add_par("Creating sample word document", style = "heading 1") %>%
body_add_par("", style = "Normal") #add space, similar to `enter` next line in word
print(word_doc, target = "sample_doc.docx")
```
- add a paragraph and style it in the created document
This formatting will involve several functions and steps:
i) Create the paragraph, store it under an object 
ii) Create the formatting object using `ftext()`. Use `fp_text()` within `ftext()` to format as appropriate. Make the formatted object a paragraph object using `fpar()`. 
iii) Add the resulting formatted paragraph to the already created `.docx` object using `body_add()` 

```r
first_paragraph <- "The `officer` package in R is a powerful tool for creating and manipulating Microsoft Word and PowerPoint documents programmatically. It provides a way to generate reports, presentations, and other documents with dynamic content, formatting, and styling. The package enables you to automate the process of document creation, making it easier to generate reproducible reports, documents, and presentations directly from your R code."

my_text_style <- ftext(first_paragraph, 
fp_text(font.family = "Maiandra GD", 
font.size = 14, color = "blue", bold = F, 
italic = T, underline = F)) %>%
fpar()

word_doc <- word_doc %>% 
body_add(my_text_style 
,style = "Normal"
,pos = "after") %>% 
body_add_par("")
print(word_doc, target = "sample_doc.docx")
```

```r
library(kableExtra)
library(flextable)

data <- iris %>% data.frame()

sample_table <- 
data %>% 
group_by(Species) %>% 
summarise(`sepal length` = mean(Sepal.Length, na.rm = T),
`sepal width` = mean(Sepal.Width, na.rm = T),
`petal length` = mean(Petal.Length, na.rm = T),
`petal width` = mean(Petal.Width, na.rm = T)) %>% 
flextable()

word_doc <- word_doc %>% body_add_flextable(sample_table)

print(word_doc, target = "sample_doc.docx")


```
