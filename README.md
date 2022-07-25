# The `xlsb2xlsx` package
The Python package (wrapper) to convert .xlsb to .xlsx files.

**WARNING**: `xlsb2xlsx` requires Python 3.9.X or **later**.

## Purpose
An alternative to `soffice` that specifically converts `.xlsb` files to `.xlsx`. 

**NOTE**: This wrapper is built on top of the [Aspose.Cells project](https://pypi.org/project/aspose-cells/) which is proprietary.

As a result of being built on top of Aspose.Cells, **each converted `.xlsx` file will contain an additional sheet named `Evaluation Warning` at the end of the `.xlsx` file after the conversion is complete**. This additional sheet reminds the user that they are using a free version of Aspose's software which should be used for "Evaluation Only". The copyrights of Aspose.Cells belong to Apose Pty Ltd and this wrapper exists only for additional ease of use.

## How to use `xlsb2xlsx`

To convert all `.xlsb` or `.XLSB` files in a directory, run:

```python -m xlsb2xlsx /path_to/directory```

To convert all `.xlsb` or `.XLSB` files in a directory *recursively*, run:

```python -m xlsb2xlsx /path_to/directory --recursive```

(The `--recursive` argument can also be shortened to `-r`). The `--recursive` flag specifies whether or not to convert `.xlsb` files in directories nested recursively within the `/path_to/directory` listed.