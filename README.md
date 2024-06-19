# Ask-LLM
This Excel addin adds two new Excel functions: ASK() and INTERPRET(). The purpose of these functions is to integrate generative AI into Excel locally via [Ollama](https://ollama.com/).

https://github.com/suncloudsmoon/Ask-LLM/assets/34616349/5b542001-402c-491f-9e15-6b0c7d4d9430
# Instructions
See `Instructions.txt`

[Download Ask-LLM (v1.0)](https://github.com/suncloudsmoon/Ask-LLM/releases/download/1.0/Ask-LLM-Addin-Files.zip)
# API
* ASK([prompt])
  * Ask a question to the LLM and get a response back in the same cell.
* INTERPRET([prompt],[data,...])
  * Feed data in the form of excel cells with the help of a prompt.
# Supported Platforms
This software only supports **Windows 7+** due to the use of .NET framework and heavy Windows related dependencies. If you want to port the software onto a different platform, feel free to fork this project.
# Credits
* [Excel-DNA](https://github.com/Excel-DNA)
* [Ollama](https://github.com/ollama/ollama)
