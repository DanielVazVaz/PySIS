from setuptools import setup

with open("README.md", "r", encoding="utf8", errors='ignore') as fh:
    long_description = fh.read()

setup(name='pysis',
      version='0.1.2',
      description='Abstract layer over Aspen HYSYS using Python',
      url='https://github.com/DanielVazVaz/PySIS',
      author='Daniel Vázquez Vázquez',
      author_email='daniel.vazquez@iqs.url.edu',
      license='MIT',
      packages=['pysis'],
      install_requires=['pywin32>=225'],
      extras_require = {
          "dev": [
              "build",
              "twine",
              "sphinx",
              "sphinx_rtd_theme",
              "check-manifest",
          ],
      },
      long_description=long_description,
      long_description_content_type="text/markdown",
      classifiers=[
              'Development Status :: 2 - Pre-Alpha',
              'License :: OSI Approved :: MIT License',
              'Programming Language :: Python :: 3 :: Only',
              'Topic :: Scientific/Engineering',
              'Topic :: Scientific/Engineering :: Mathematics'
          ],
)