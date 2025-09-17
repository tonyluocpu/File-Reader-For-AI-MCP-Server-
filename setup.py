#!/usr/bin/env python3
"""
Setup script for Word Reader MCP Server
"""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

with open("requirements.txt", "r", encoding="utf-8") as fh:
    requirements = [line.strip() for line in fh if line.strip() and not line.startswith("#")]

setup(
    name="word-reader-mcp-server",
    version="1.0.0",
    author="Your Name",
    author_email="your.email@example.com",
    description="A Model Context Protocol server for reading Microsoft Word documents",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/word-reader-mcp-server",
    py_modules=["word_reader_mcp_server"],
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Text Processing :: Markup",
        "Topic :: Office/Business",
    ],
    python_requires=">=3.8",
    install_requires=requirements,
    entry_points={
        "console_scripts": [
            "word-reader-mcp=word_reader_mcp_server:main",
        ],
    },
    keywords="mcp, model-context-protocol, word, docx, document, ai, assistant",
    project_urls={
        "Bug Reports": "https://github.com/yourusername/word-reader-mcp-server/issues",
        "Source": "https://github.com/yourusername/word-reader-mcp-server",
        "Documentation": "https://github.com/yourusername/word-reader-mcp-server#readme",
    },
)
