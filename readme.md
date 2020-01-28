# CoverLettterGenerator

generate multiple cover letters for multiple companies from one single word document template.

## How to run

the app is a windows forms application coded in c# in visual studio 2017.

1. clone the app or download as zip
2. open your visual studio
3. File -> open -> projet/solution
4. go to the project directory and open `coverLetterGenerator.sln`
   > [!IMPORTANT]
   > before you run the app copy the file `coverLetterGenerator\CoverTemp.docx` to `coverLetterGenerator\bin\Debug`
   > the app uses the path of the .exe file to look for the word document

## How the app works

the app read the `CoverLetter.docx` and replace the text as you will seen in the code named `<fullname>`, `<position>`, `<company>`, `<finish>`
with the text form the datagridview. the `<finish>` get replaced with "your sincerely" or "your faithfully" based on if you type "y" on the finish column or not.
and generate cover letter based on how many rows you filled on the interface.

read the comment on the code to well understand it
