# VBA Matrix and Double Utility Classes

This project provides a set of VBA classes and modules for matrix operations and encapsulated double values, suitable for use in Excel or other VBA environments. It is designed for educational, scientific, and engineering applications where matrix computations are required.

## Features

- **Algelin.cls**: Module containing matrix operations routines. Supports:
  - Random matrix generation (integer/double, symmetric/banded)
  - Matrix arithmetic (add, subtract, multiply)
  - LU decomposition and linear system solving
  - Eigenvalue and Eigenvector approximations
  - Utility functions (transpose, trace, symmetry check, etc.)

- **Utilities.bas**: Helper functions for generating random numbers (integer and double) within specified bounds.

- **Shee1.doccls**: Example/test module demonstrating usage of Linear System and Eigvec/Eigvalues routines.

- **EngLogger.cls**: Logger class for printing errors, warnings, and informational messages.

## Usage

1. **Import the files** into your VBA project (Excel, Access, etc.).
2. Use the subroutines in `Sheet1.doccls` as a starting point to play around and get used to the methods.
3. Extend or modify as needed for your application.

## License
MIT License

## Author
- https://github.com/lz0-0zl

---
Feel free to open issues or contribute improvements!
