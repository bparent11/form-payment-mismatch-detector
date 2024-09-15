# form-payment-mismatch-detector
Python program to make verifications about a list of people who filled out a form and another list of people who paid a license fee. We want to see who are the prompt payers, the delinquents, and who has paid without filling out the form.

In this repo, you can try my code located in the directory named "FakeData_Try" (I simplified the test here), the fact is that the initial code is really used by an association and I could'nt share this data on this repo caused by confidentiality issues obviously (that's why I don't tell you to test my code with the file "verif_cotisants_psl.ipynb").
I generated a list of "fake" names and surnames which are actually names and surnames of fantastic characters, thanks to chatGPT by the way.

So you have the xlsx file of the "cotisants", people who have paid a fee to get their licence. Then, the xlsx file of the people who have filled out the form to participate to an event (reserved for members of the association (those who have paid the fee)).
You can run the code, and you'll get the list of the three kind of people.

Bonus : You can display the results in a html page, you'll need to export your data is json format and use javascript as I did with these files : verif_cotisants_psl.ipynb, script.js and index.html.