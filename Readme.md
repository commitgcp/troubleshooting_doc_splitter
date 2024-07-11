# Troubleshooting Document Splitter

Takes all .doc or .docx files in the input directory, splits them by header, and calls an LLM to describe any images (diagrams or tables), and adds the LLM's response to the resulting split documents. Then converts the documents to PDF.

The output is saved in directory output/final_output