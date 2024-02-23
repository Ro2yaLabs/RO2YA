from setuptools import setup, find_packages
import codecs
import os

here = os.path.abspath(os.path.dirname(__file__))

with codecs.open(os.path.join(here, "README.md"), encoding="utf-8") as fh:
    long_description = "\n" + fh.read()

VERSION = '0.0.1'
DESCRIPTION = 'Multilingual AI package: Edu-Tech, Automation, Sound, 2D/3D Lip Sync. Low-code, free, real-time for individuals and enterprises.'
LONG_DESCRIPTION = 'An AI package written in Python and C++ that facilitates the accurate Multilingual & Zero-Cost usage of the best building blocks generation algorithms, edu-tech recommendation systems in various scales, complex automation processes, sound generation & transcription processes, and even the 2d & 3d lipsyncing technologies. All in a low code package with free arguments and a real-time implementation for individuals and enterprises.'

# Setting up
setup(
    name="ro2ya",
    version=VERSION,
    author="Ro2ya Labs, MGE",
    author_email="<ro2yaproject@gmail.com>",
    description=DESCRIPTION,
    long_description_content_type="text/markdown",
    long_description=LONG_DESCRIPTION,
    packages=find_packages(),
    install_requires=['opencv-python', 'pyautogui', 'pyaudio', 'langchain', 'python-docx', 'numpy', 'pytorch', 'tensorflow', 'ffmpeg'],
    keywords=['python', 'video', 'stream', 'video stream', 'camera stream', 'sockets'],
    classifiers=[
        "Development Status :: 1 - Planning",
        "Intended Audience :: Developers",
        "Programming Language :: Python :: 3",
        "Programming Language :: C++",
        "Operating System :: Unix",
        "Operating System :: MacOS :: MacOS X",
        "Operating System :: Microsoft :: Windows",
    ]
)
