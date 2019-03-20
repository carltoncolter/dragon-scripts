
# dragon-scripts
This project/repository is a meant to be a collection of Dragon NaturallySpeaking Scripts and DragonFly NatLink Scripts that simplify voice control.  Most of the code in this repository is visual basic scripts or python.  If/when other languanges are leveraged, I will do my best to keep this page updated with the basics.

**NO WARRANTY**
---
>THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
---
>**What does that mean?**  All code and information in this repository is provided **"AS IS"**.  You may use it at your own risk.  The people who wrote or contributed the code to this repository are not liable for any consequences from your use of material in this repository.  You should **ALWAYS** take the necessary steps to ensure you have a good backup and and a planned restore process before using anything in this repository.
---

## Dragon NaturallySpeaking Scripts
All Dragon NaturallySpeaking Scripts include both the dat file. If it leverages a vb-script then the vb-script is included as a .vbs file.

## Natlink Dragonfly Scripts
These scripts are written in python and are leveraged by Dragon Naturally Speaking integrated with Natlink and Dragonfly.

## Installation
### Installing Natlink and Dragonfly
> People recommend avoiding 64-bit installations of Java, Python, Natlink, and Dragonfly.  I tried going the 64-bit route and couldn't get it to work.  I also tried updating everythin gto using Python 3 and haven't gotten that to work.  It is a major undertaking, which will likely be a project unto itself that will take a lot longer.
> [Poppe1219](https://github.com/poppe1219) posted in the [dictation toolbox dragonfly-scripts](https://github.com/dictation-toolbox/dragonfly-scripts) repository that he had problems using Eclipse with 64-bit due to the Java virtual machine.  He was having an issue with double typing of the first character or loss of the first character, which he fixed by installing the Java virtual machine 32-bit version instead.  So just stay with the 32-bit installs for now.

*NOTE: If you have multiple Python versions, make sure you install all packages into the correct python version.*

#### Installation Steps:

1. Install Dragon NaturallySpeaking
2. Install [Python 2.7 32-bit version](https://www.python.org/downloads/)
   * Python 2.7 and 3 are pretty different and NatLink doesn't currently support 3.
3. Update your environment path to include the python bin folder.
4. Upgrade pip by running: *python -m pip install --upgrade pip*
5. Install Python Packages (these are part of the steps to install natlink):
   * pip install wxpython
   * pip install pywin32
   * pip install six
6. Install [NatLink](https://sourceforge.net/projects/natlink/) ([Instructions](https://qh.antenna.nl/unimacro/installation/installation.html))
7. Install [DragonFly](https://github.com/dictation-toolbox/dragonfly)
   * pip install dragonfly2
