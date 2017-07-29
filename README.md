# email-mistakes
An outlook visual basic macro to make sure you are not mixing domains when sending emails

Based on https://www.slipstick.com/how-to-outlook/prevent-sending-messages-to-wrong-email-address/

The macro iterates over the recipeients of the email and counts the number of "foreign" domains included in the list. If there is more than one, it warns the user with a message.

## Installation instructions
This macro needs to be added to ThisOutlookSession to work.
You also need to enable the right level of security to allow the macro to run whenever you send a new message.
Presee alt+F11 in outlook to get to the Marco Editor.
