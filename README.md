# ElectroPy

This is a Python program that reads an excel file with information from eletrophysiology recordings on brain slices, analyses it and exports the required parameters in an accessible way. I made it for my girlfriend, who is currently a Neuroscience Masters’ student. She is studying the role of inflammation on an epilepsy model, using hippocampal slices to assess both molecular and electrophysiology hallmarks in the progression of the disease. She submits those slices to a 40-minute protocol where a specific program detects spontaneous brain activity and then returns those oscillations in data: time and amplitude. She asked for my help to optimise that data’s treatment because the method she used was archaic and very time consuming. From the results the program returns, she needed to apply certain temporal conditions to find bursts of activity, on top of which she needed to extract the number of bursts, the duration and frequency of each burst, the number of events per burst and the peak and average amplitude of each burst. All of this for a single sample, and she has around 12 samples per week.

Please keep in mind this was my first time trying to create something like this so I am positive that there are more efficient ways to simulate what I did. I still have a lot to develop in my Python skills, but the goal with this program was to provide, with what I know today, an user-friendly solution that could simplify my girlfriend's tasks of analysing each sample's data.

## 📁 Documentation

- [Python analysis](https://github.com/MPCaloba/EletroPy/blob/main/Program.py)
- [Sample datasets](https://github.com/MPCaloba/EletroPy/blob/main/Sample%20Files.zip)

## 🔗 Links
- [Blog Article](https://medium.com/@codex-ventures/eletropy-electrophysiology-analysis-program-5230ae47e352)
- [Medium](https://medium.com/@codex-ventures)
