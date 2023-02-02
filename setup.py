import setuptools

with open("README.md", "r") as fh:
    description = fh.read()

setuptools.setup(
    name="pywin-query",
    version="0.0.1",
    author="KSokhal",
    author_email="ksokhal99@gmail.com",
    packages=["pywin_query"],
    description="An convenience wrapper that allows you to use Windows' File Search functionality.",
    long_description=description,
    long_description_content_type="text/markdown",
    url="https://github.com/KSokhal/pywin-query",
    license="MIT",
    python_requires=">=3.8",
    install_requires=["pywin32"],
)
