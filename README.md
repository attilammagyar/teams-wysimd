WYSIMD (for MS Teams) - 0.11 ALPHA
==================================

<img src="https://raw.githubusercontent.com/attilammagyar/teams-wysimd/master/teams-wysimd.gif" alt="screen demo video showing the editor with instant formatting turned off" />

[Greasemonkey](https://www.greasespot.net/) /
[Tampermonkey](https://www.tampermonkey.net/) user script to turn off the
[WYSIWYG](https://en.wikipedia.org/wiki/WYSIWYG) editor in
[Microsoft Teams](https://teams.microsoft.com/) and turn it into a raw
[Markdown](https://en.wikipedia.org/wiki/Markdown) editor with a syntax
inspired by the one in
[Slack](https://slack.com/intl/en-hu/help/articles/202288908-Format-your-messages).

WARNING
-------

This script is probably very fragile, it might break in unexpected ways with
(or without) any update to Teams, it might have browser compatibility issues,
and I'm not sure how much time I can afford to keep it working or how soon I
can get it to work again if it breaks. Also, I don't use Teams very
extensively, so I might not even notice if this script interferes with some
feature or message extension that might be essential for some people. Use it at
your own risk.

Installing
----------

1. If you use Firefox, install the [Greasemonkey](https://www.greasespot.net/)
   extension, or if you use Chrome, install the
   [Tampermonkey](https://www.tampermonkey.net/) extension.
2. Add [teams-wysimd.user.js](https://github.com/attilammagyar/teams-wysimd/raw/main/teams-wysimd.user.js)
   as a new user script.

Why?
----

After having been writing raw Markdown for many years, I just couldn't get used
to Slack's WYSIWYG editor, and now I can't get used to the one in Teams. While
[Slack allows users to turn off the WYSIWYG mode](https://twitter.com/SlackHQ/status/1201955273667158023),
unfortunately [Teams doesn't](https://microsoftteams.uservoice.com/forums/555103-public/suggestions/20588818-raw-markdown-editor-for-messaging).
So if one really wants to write raw Markdown without instant rendering and
without annoying bugs (like when backticks are not turned into monospace if
they immediately follow a parenthesis), one has to either use external editors
to compose messages and then copy&paste them into Teams, or one has to resort
to other kinds of trickery. Hence, WYSIMD (What You See Is MarkDown) was born.

Features
--------

 * **ON/OFF switch**: (bottom right corner) while WYSIMD can be useful in a
   chat window or for team posts (where the Enter key sends the message), it
   can mess up the editor's behaviour in the Wiki tab and in the Calendar, and
   I couldn't figure out yet how to detect those cases. But even if I did, one
   might still want to temporarily turn off this thing for one reason or
   another. And if the switch gets in the way, just keep the mouse over it, and
   after a few seconds, it will disappear for a moment.

 * **Custom emojis**: though it doesn't allow
   [custom reaction emojis](https://microsoftteams.uservoice.com/forums/555103-public/suggestions/16934329-allow-adding-custom-emoji-memes-gifs-reactions),
   at least you can put them in your chat messages. Since I don't want to host
   copyrighted content on GitHub, I won't include a
   [Hide the Pain Harold](https://en.wikipedia.org/wiki/Andr%C3%A1s_Arat%C3%B3)
   face or a sarcastic slow-clap animation, but you can put your own favorites
   in the script.

 * **Non-intrusive**: it tries not to interfere with all the whistles and bells
   Teams can do, like mentions, badges, cards, etc. And you can still use the
   usual keyboard shortcuts like `Ctrl+B` or `Ctrl+I` to format any selected
   text, etc. If you go back to edit an already formatted message, WYSIMD will
   try to keep the already formatted parts.

Quirks
------

 * **Wiki, Calendar**: those editors where the `Enter` key has no special
   functionality, will not behave nicely with WYSIMD: e.g. when you hit `Enter`
   to add a new line, the cursor will jump to the top.

 * **Auto-resizing, emoji selector**: in order to prevent Teams from doing its
   premature formatting, WYSIMD needs to temporarily disable some event
   handlers of the editor widget when a Markdown key on the keyboard is
   pressed. Unfortunately, this might interfere sometimes with Teams'
   capability to grow the editor widget as its content gets longer, or to
   remove some of the characters from the editor that are typed while the emoji
   selector is visible.

 * **Lists**: though WYSIMD disables the automatic formatting of ordered and
   unordered lists (lines which start with "`1. `" or "`- `"), it doesn't
   provide a Markdown alternative, so you have to type your bulletpoints and
   numbers one by one manually. However, this has the advantage that when the
   resulting formatted text is copied and pasted into a plain text document,
   the numbers and the bulletpoints won't be replaced with tabs.

Syntax
------

 * **Bold**

       *bold text*

 * **Italic**

       _italic text_

 * **Strikethrough**

       ~strikethrough text~

 * **Inline code**

       `inline code`

 * **Blockquote**

       > start a line with a greater-than sign

 * **Code block**

       3 backticks will start a
       code block anywhere: ```
       if (1 + 1 === 2) {
           console.log("Hello World");
       }``` another 3 backticks close it

 * **Custom emoji**

       here's a loading spinner: :spinner:

Custom Emojis
-------------

WYSIMD comes with 3 built-in emojis:

 * `:pf:` is a [pinched-fingers](https://openmoji.org/library/#emoji=1F90C)
   emoji from [OpenMoji](https://openmoji.org/) - the open-source emoji and
   icon project. License:
   [CC BY-SA 4.0](https://creativecommons.org/licenses/by-sa/4.0/#)

   Note that withot WYSIMD, Teams would convert it into a
   [face-with-tongue](https://emojipedia.org/face-with-tongue/) emoji as soon
   as you type `:p`.

 * `:pm:` and `:spinner:` are loading-animations generated with
   [Ajaxload](http://www.ajaxload.info/).
   License: [WTFPL](http://www.wtfpl.net/)

Creating Custom Emojis
----------------------

Near the top of the script, you find this code:

    custom_emojis = {
        "spinner": {
            "src": "data:image/gif;base64,R0l...",
            "width": 16,
            "height": 16
        },
        "pm": {
            "src": "data:image/gif;base64,R0l...",
            "width": 24,
            "height": 24
        },
        "pf": {
            "src": "data:image/gif;base64,R0l...",
            "width": 20,
            "height": 20
        }
    };

You can add your own custom emojis here by converting them into
[Data URL](https://en.wikipedia.org/wiki/Data_URI_scheme) strings, e.g. using
an [online data URL image converter](https://ezgif.com/image-to-datauri).

Though some message extensions that seem to be built into Teams can come with
several kilobytes of HTML and SVG data, I don't recommend putting huge images
and animations as data URLs into the script and using them excessively. It's a
good idea to resize your custom emojis to about 20x20 pixels (Teams' own emojis
are this size when they are embedded in text), and to avoid animations or at
least to lower their framerate, and to reduce the color palette of GIFs. Even
just a few big custom emojis can make a message so big that Teams won't let
you send it.
