# [xlsb2xlsx](https://github.com/brianthomaslewis/xlsb2xlsx)

## Purpose
An alternative to `soffice` that specifically converts `.xlsb` files to `.xlsx` files. 

# How to use `xlsb2xlsx`

To convert all `.xlsb` files in a directory, run:

```python -m xlsb2xlsx /filepath```

To convert all `.xlsb` files in a directory *recursively*, run:

```python -m xlsb2xlsx /filepath --recursive```

OR

```python -m xlsb2xlsx /filepath -r```

The `--recursive` flag specifies whether files in nested directories are also converted.

## Fine Print

This wrapper is built on top of the [Aspose.Cells project](https://pypi.org/project/aspose-cells/) which is proprietary.

As a result of being built on top of Aspose.Cells, **each converted `.xlsx` file will contain an additional sheet named `Evaluation Warning` at the end of the `.xlsx` file after the conversion is complete**. This additional sheet reminds the user that they are using a free version of Aspose's software which should be used for "Evaluation Only". The copyrights of Aspose.Cells belong to Apose Pty Ltd and this wrapper exists only for additional ease of use.

#

