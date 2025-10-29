---
title: "Markdown to DOCX Converter Demo"
author: "990aa"
date: "October 29, 2025"
---

# Introduction

This document demonstrates **all features** of the MD to DOCX converter including LaTeX math, code highlighting, and tables.

## Mathematical Equations

### Inline Math

The equation \( \frac{1 - x}{2} = \frac{y - 1}{3} = \frac{z}{1} \) represents a line in 3D space.

The domain is \(\mathbb{R} - \{10\}\) and the magnitude is \(| \overrightarrow{a} | = 2\).

### Block Equations

Matrix representation:

\[
\begin{bmatrix} 
2 & -1 & 1 \\ 
\lambda & 2 & 0 \\ 
1 & -2 & 3 
\end{bmatrix}
\]

Quadratic formula:

\[
x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}
\]

## Code Examples

### Python Code

```
def factorial(n):
"""Calculate factorial of n"""
if n <= 1:
return 1
return n * factorial(n - 1)

print(factorial(5)) # Output: 120
```


### JavaScript Code

```
const fetchData = async (url) => {
const response = await fetch(url);
return response.json();
};
```


## Tables

| Feature | Supported | Notes |
|---------|-----------|-------|
| LaTeX Math | ✅ Yes | Full support with Word equation editor |
| Code Highlighting | ✅ Yes | Multiple syntax styles available |
| Tables | ✅ Yes | Converted to native Word tables |
| Images | ✅ Yes | Embedded automatically |

## Lists

### Ordered List

1. Initialize project with `uv init`
2. Install dependencies with `uv add`
3. Run converter with `uv run`

### Unordered List

- **Fast**: uv is written in Rust
- **Simple**: No external installations needed (except Pandoc)
- **Flexible**: Supports custom styling via reference documents

