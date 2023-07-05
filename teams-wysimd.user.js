// ==UserScript==
// @name            WYSIMD (for MS Teams)
// @description     Turn the WYSIWYG editor in MS Teams into a Slack-inspired Markdown editor
// @version         0.1
// @grant           none
// @match           https://teams.microsoft.com/*
// ==/UserScript==
(function () {

function WYSIMD()
{
    var md_chars = "`*~_-1. ",
        self_closing_tags = /^(area|base|br|col|embed|hr|img|input|link|meta|param|source|track|wbr|command|keygen|menuitem)$/,
        log_prefix = "WYSIMD: ",
        log_debug,
        log_trace,
        debugging = false,
        tracing = false,
        _window = window,
        failures = 0,
        stop_change_events = 0,
        previous_key = "",
        custom_emojis,
        is_enabled = true,
        custom_emoji_initials;

    custom_emojis = {
        "spinner": {
            "src": "data:image/gif;base64,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",
            "width": 16,
            "height": 16
        },
        "pm": {
            "src": "data:image/gif;base64,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",
            "width": 24,
            "height": 24
        },
        "pf": {
            "src": "data:image/gif;base64,R0lGODlhQABAAIQQAAACAA4NAhYTAhsZAyEgBTEuB0pEDFhTDm1mEY+EF7KnHM6/H+HQI+/gI/rqKf/yI////////////////////////////////////////////////////////////////yH5BAEKABAALAAAAABAAEAAAAX+ICSOZGmeaKqubOuqAiAPb22rcg4Ed2/rAEEhBvAZV0AZ4dEI7I7QUo6QUBhkjAZBFjU6ZUSC4+FIyBZabtf2TTYcjYZCpnBci+tX+5AYgOsODnMACQ4IanktOQx2OgGFDgsyCQ9meIlIO1lbSYwNMgiRiJgoRIWHAAQKD4OQMgYODKOkJkpMYGSxoGNOsA9gtCiThmeBDbIAoQ9bBQ8PX8G1Mm9+AW+ByL5XAs4FT9Ekww0xAoHGMgWBqM534CMy1nC95nDoY5ZkqO4iMgfG89ccOBEjaFqZWbRymGPgxB+9HQQlAagzyF0OSPIAwKIXYwA2OgUvBQNDjyEAh4H+HnjzeGyXRB4jsdBrotHcmysBsDXUhXDNl1Azd9I7EOQNzXQOctAKQ+9fTXqorpEbo5TUoqY8gZozY+3NVJU9oeRYgLURAEb0LF3bEo+oyCicCM0MJHEjvUHXcL6xlIfmRDhYiQYABNUgsbMhj+wwcA1BAcLGzgFo1vQTAI+BBDNCdqPN1bJN75ANeJCQUSfxLNdIooQ04KYmMUdOGiSXJUg0BbRoE6BAggXOcr2eHcgS5OLDApE0FgMmjough2O1jLSpHwDXbgd0ovu5RrSVw9s8jNFcq5TXA1ruLgykdAZZojvNOV7gDvPJ6YblV2w6agO/ZcHAAgo4FsAXx4X+5BBNAYBnVgrQhWcZazm0wRhW3mD3xiAo1bOfLde41gACAnjmyAC+5RKiREAtI0NTg7BXwhfguTbTgArkuAB81whnU4ZkhaRVIO20V110NoJmoxkoMSCAAJ5I5pw05c0l35WR+QiaWyl8IZxRWCpJXJiBENHlZPKBmSaZ0UlEQ3sLIinmdGyaw8lzQ16pHpZq2njNS8/hNtueayappwNc4smmoWGGWBk8Kwg1pnR0Wukoow9mQpB4YzpK6ZzjIfPWCUTIOema9U2aKAuTYMqplXUCqkhRr4lY56Sk3bSfCQjW6WmqwAZkiYwsfPGrp4Sm6qqHo0b6yq2wlmVjhm9o1vCFXaDiGm1h39wADwLIVpotcag0+0IOUN7q6hsLxBXFFwEY4KCtVxKIgAEE8LbGDjn4psCOAw64wMA5KpBAAggcgK+JFZprhCMUQsywI1OSciC/qKHGLxAHdrwPCgeWiFq1H5ccRQgAOw==",
            "width": 20,
            "height": 20
        }
    };

    function noop()
    {
    }

    function log_info(msg)
    {
        console.log(log_prefix + String(msg));
    }

    log_debug = debugging ? log_info : noop;
    log_trace = tracing ? log_info : noop;

    function md_to_html(md)
    {
        /**
         * This spaghetty tries to take the text which is visible on the screen
         * inside the WYSIWYG editor, and interpret just the visible parts as a
         * subset of a Slack + Markdown monstrosity and convert it into HTML,
         * without messing up the existing HTML inside the editor, like
         * mentions, badges, already formatted text, etc.
         *
         * However, the job is little bit more complicated than it may seem at
         * first. For example, when the editor contains something like this:
         *
         *     markdown with *strong* and _italic_ and *_strong + italic_*,
         *     maybe ~strikethrough~ and `inline code`
         *     then a code block
         *     ```var arr = [];
         *     for (key in obj) {
         *         if (obj.hasOwnProperty(key)) {
         *         }
         *     }```
         *     and some more text
         *
         * then the DOM tree inside the editor can look like this (indentation
         * added for readability):
         *
         *     <div>
         *         markdown with *strong* and _italic_ and *_strong + italic_*,
         *         maybe ~strikethrough~ and `inline code`
         *         <br>
         *     </div>
         *     <div>
         *         then we have some ```
         *         <br>
         *     </div>
         *     <div itemprop="copy-paste-block" class="copy-paste-block">
         *         <div>
         *             var arr = [];
         *             <br>
         *         </div>
         *         <div>
         *             for (key in obj) {\u200b
         *             <br>
         *             &nbsp;&nbsp;&nbsp; if (obj.hasOwnProperty(key)) {\u200b
         *             <br>
         *             &nbsp;&nbsp;&nbsp; }\u200b
         *             <br>
         *             }\u200b```
         *             <br>
         *         </div>
         *     </div>
         *     <div>
         *         and some more text
         *         <br>
         *     </div>
         *
         * (Or some variation of that, depending on how the text was typed,
         * copy&pasted, etc.)
         *
         * Most of the markdown formatting is simple and easy inside a single
         * line, but fenced code blocks have to break up the line if they are
         * in the middle of it, and a code block may span across multiple
         * `<div>` elements. We also need to keep track of whether a line is
         * inside a code block or not, because we don't want to have formatting
         * accidents like emojis inside code. Also, we don't want to do any
         * formatting inside HTML elements other than those bare and copy&paste
         * `<div>` nodes, because we don't want to mess up mentions, badges,
         * etc. especially in messages that were already sent, and just opened
         * again for editing like fixing typos, etc.
         *
         * Long story short, trying to keep the markdown somewhat compatible
         * with Slack's version seemed like a good idea at first, but looking
         * at this code, I'm not so sure anymore...
         */

        function is_bare_div(obj)
        {
            return (
                String(obj.nodeName).toLowerCase() === "div"
                && obj.attributes
                && obj.attributes.length === 0
            );
        }

        function has_line_end(text)
        {
            return text.match(/([\r\n]|(<br\s*(\s+[^>]*)*\/*>))$/is); /* <br type="_moz"> */
        }

        function remove_bare_divs(text)
        {
            var result = "",
                node_iter, node_copier, child_node, node_name, inner_html, i, l;

            if (!text) {
                return "";
            }

            log_trace("remove_bare_divs");
            log_trace(JSON.stringify(text));

            node_iter = document.createElement("div");
            node_iter.innerHTML = text;

            for (i = 0, l = node_iter.childNodes.length; i < l; ++i) {
                child_node = node_iter.childNodes[i];

                if (child_node.nodeType === 1) {
                    inner_html = remove_bare_divs(child_node.innerHTML)

                    if (is_bare_div(child_node)) {
                        if (result && !has_line_end(result)) {
                            result += "\n";
                        }

                        result += inner_html;

                        if (!has_line_end(inner_html)) {
                            result += "\n";
                        }
                    } else {
                        child_node.innerHTML = inner_html;
                        result += child_node.outerHTML;
                    }
                } else {
                    node_copier = document.createElement("div");
                    node_copier.appendChild(child_node.cloneNode(true));
                    result += node_copier.innerHTML;
                    node_copier = null;
                }
            }

            return result;
        }

        function add_custom_emojis(text)
        {
            var name, emoji;

            for (name in custom_emojis) {
                if (!custom_emojis.hasOwnProperty(name)) {
                    continue;
                }

                emoji = custom_emojis[name];

                if (!emoji.hasOwnProperty("img")) {
                    emoji["img"] = (
                        "<img src=\"" + emoji["src"] + "\""
                        + " alt=\":" + name + ":\""
                        + " title=\":" + name + ":\""
                        + " width=\"" + String(emoji["width"]) + "\""
                        + " height=\"" + String(emoji["height"]) + "\" />"
                    );
                }

                text = text.replace(":" + name + ":", emoji["img"]);
            }

            return text;
        }

        function format_line(line)
        {
            var parts = [String(line)],
                html = [],
                in_code = false,
                can_close = false,
                CODE_DELIM = 1,
                part, suffix, i;

            while (parts.length) {
                part = parts.shift();

                if (part === CODE_DELIM) {
                    if (in_code) {
                        html.push("</code>");
                        in_code = false;
                    } else if (can_close) {
                        html.push("<code>");
                        in_code = true;
                    } else {
                        html.push("`");
                    }
                } else {
                    i = part.indexOf("`");

                    if (i > -1 && part[i + 1] !== "`") {
                        can_close = false;

                        if (i < part.length - 1) {
                            suffix = part.substr(i + 1, part.length);
                            parts.unshift(suffix);
                            can_close = suffix.indexOf("`") > -1;
                        }

                        parts.unshift(CODE_DELIM);

                        if (i > 0) {
                            parts.unshift(part.substr(0, i));
                        }
                    } else if (in_code) {
                        html.push(part);
                    } else {
                        html.push(
                            add_custom_emojis(
                                part.replace(/\*([^*\n\r]+)\*/g, "<strong>$1</strong>")
                                    .replace(/_([^_\n\r]+)_/g, "<em>$1</em>")
                                    .replace(/~([^~\n\r]+)~/g, "<del>$1</del>")
                                    .replace(
                                        /(^|\n)&gt; (.*)$/,
                                        "$1<blockquote>$2</blockquote>"
                                    )
                            )
                        );
                    }
                }
            }

            return html.join("");
        }

        var remaining = remove_bare_divs(
                String(md).replace(/\r\n/g, "\n").replace(/\r/g, "\n")
            ),
            lines = [ remaining ],
            html = [],
            in_code_block = false,
            open_tags = [],
            line, parts, suffix, tag, closing, attrs, tag_name, self_closing, i, l;

        log_trace("converting markdown to HTML");
        log_trace(JSON.stringify(lines));

        while (lines.length) {
            line = lines.shift();

            log_trace("processing line:");
            log_trace(JSON.stringify(line));
            log_trace("remaining:");
            log_trace(JSON.stringify(remaining));
            log_trace("open tags:");
            log_trace(JSON.stringify(open_tags));

            if (typeof(line) === "string") {
                if ((i = line.indexOf("<")) > -1) {
                    if (i > 0) {
                        lines.unshift(line.substr(i));
                        lines.unshift(line.substr(0, i));
                        continue;
                    }

                    tag = line.match(
                        /^<(\/*[\s\/]*)([^>\s]*)(\s*[^\s=>]*(=(("[^"]*")|([^\s\/>]*))?)?)*\s*(\/*)\s*>/m
                        // 1 closing   2 tag    3 attributes                                 8 self-closure
                    );

                    if (tag) {
                        closing = tag[1];
                        tag_name = String(tag[2]).toLowerCase();
                        attrs = tag[3];
                        self_closing = (
                            Boolean(tag[8])
                            || Boolean(tag_name.match(self_closing_tags))
                            || tag_name[0] === "!"
                        );
                        tag = tag[0];

                        if (tag_name !== "br") {
                            log_trace(
                                "found HTML:"
                                + " tag_name=" + tag_name
                                + " attrs=" + attrs
                                + " closing=" + closing
                                + " self_closing=" + self_closing
                                + " tag=" + tag
                            );

                            if (line.length > tag.length) {
                                lines.unshift(line.substr(tag.length))
                            }

                            lines.unshift({"type": "raw", "raw": tag});

                            if (!self_closing) {
                                if (closing) {
                                    if ((i = open_tags.lastIndexOf(tag_name)) > -1) {
                                        open_tags.splice(i, 1);
                                    } else if (
                                        tag_name === "div"
                                        && (i = open_tags.lastIndexOf("#copy-paste-block")) > -1
                                    ) {
                                        open_tags.splice(i, 1);
                                    }
                                } else {
                                    if (tag_name === "div" && attrs.match(/copy-paste-block/)) {
                                        tag_name = "#copy-paste-block";
                                    }

                                    open_tags.push(tag_name);

                                    if (in_code_block && tag_name !== "#copy-paste-block") {
                                        lines.unshift({"type": "fenced-code-block"});
                                    }
                                }
                            }

                            continue;
                        }
                    }
                }

                if ((i = line.indexOf("<br>")) === 0) {
                    if (i < line.length - 4) {
                        lines.unshift(line.substr(i + 4, line.length));
                    }

                    lines.unshift({"type": "br"});

                    if (i > 0) {
                        lines.unshift(line.substr(0, i));
                    }
                } else if ((i = line.indexOf("\n")) > -1) {
                    if (i < line.length - 1) {
                        lines.unshift(line.substr(i + 1, line.length));
                    }

                    lines.unshift({"type": "newline"});

                    if (i > 0) {
                        lines.unshift(line.substr(0, i));
                    }
                } else if (
                    open_tags.length > 0
                    && (
                        open_tags.length !== 1
                        || open_tags[0] !== "#copy-paste-block"
                    )
                ) {
                    html.push(line);
                    remaining = remaining.substr(line.length);
                } else {
                    i = line.indexOf("```");

                    if (
                        (i > -1)
                        && (
                            in_code_block
                            || line.lastIndexOf("```") > i
                            || remaining.substr(line.length).indexOf("```") > -1
                        )
                    ) {
                        if (i < line.length - 3) {
                            suffix = line.substr(i + 3, line.length);
                            lines.unshift(suffix);
                        }

                        lines.unshift({"type": "fenced-code-block"});

                        if (i > 0) {
                            lines.unshift(line.substr(0, i));
                        }
                    } else if (in_code_block) {
                        html.push(line);
                        remaining = remaining.substr(line.length);
                    } else {
                        html.push(format_line(line));
                        remaining = remaining.substr(line.length);
                    }
                }
            } else if (line.type === "raw") {
                html.push(line.raw);
                remaining = remaining.substr(line.raw.length);
            } else if (line.type === "newline") {
                html.push("<br>");
                remaining = remaining.substr(1);
            } else if (line.type === "br") {
                html.push("<br>");
                remaining = remaining.substr(4);
            } else if (line.type === "fenced-code-block") {
                if (in_code_block) {
                    html.push("</code></pre><br>");
                    in_code_block = false;
                } else {
                    html.push("<br><pre><code>");
                    in_code_block = true;
                }
                remaining = remaining.substr(3);
            }
        }

        html = html.join("");

        return html;
    }

    function test_md_to_html()
    {
        var tests = [
            [
                "plain `some code` plain",
                "plain <code>some code</code> plain"
            ],
            [
                "plain ~deleted text~ plain",
                "plain <del>deleted text</del> plain"
            ],
            [
                "plain *strong text* plain",
                "plain <strong>strong text</strong> plain"
            ],
            [
                "plain _italic text_ plain",
                "plain <em>italic text</em> plain"
            ],
            [
                "plain _italic *and strong* text_ plain",
                "plain <em>italic <strong>and strong</strong> text</em> plain"
            ],
            [
                "plain *strong _and italic_ text* plain",
                "plain <strong>strong <em>and italic</em> text</strong> plain"
            ],
            [
                "plain *strong _and italic_ text* plain",
                "plain <strong>strong <em>and italic</em> text</strong> plain"
            ],
            [
                "strong *formatting\ndoes not apply across* lines",
                "strong *formatting<br>does not apply across* lines"
            ],
            [
                "italic _formatting\ndoes not apply across_ lines",
                "italic _formatting<br>does not apply across_ lines"
            ],
            [
                "strikethrough ~formatting\ndoes not apply across~ lines",
                "strikethrough ~formatting<br>does not apply across~ lines"
            ],
            [
                "code `formatting\ndoes not apply across` lines",
                "code `formatting<br>does not apply across` lines"
            ],
            [
                "<div>code `formatting<br></div><div>does not apply across` HTML lines</div>",
                "code `formatting<br>does not apply across` HTML lines<br>",
            ],
            [
                "*code `wins* over strong`",
                "*code <code>wins* over strong</code>"
            ],
            [
                "_code `wins_ over italic`",
                "_code <code>wins_ over italic</code>"
            ],
            [
                "~code `wins~ over strikethrough`",
                "~code <code>wins~ over strikethrough</code>"
            ],
            [
                "code block:```  first line\n  second line``` some more text",
                "code block:<br><pre><code>  first line<br>  second line</code></pre><br> some more text"
            ],
            [
                "first code block: ```first``` second code block: ```second``` end",
                (
                    "first code block: <br><pre><code>first</code></pre><br>"
                    + " second code block: <br><pre><code>second</code></pre><br>"
                    + " end"
                )
            ],
            [
                "`no *strong* or _italic_ or ~strikethrough~ inside code`",
                "<code>no *strong* or _italic_ or ~strikethrough~ inside code</code>"
            ],
            [
                "code blocks that can't be ``` closed are left alone",
                "code blocks that can't be ``` closed are left alone"
            ],
            [
                (
                    "<div>prefix</div><div>&gt; block quotes with *strong _and italic_ text* work</div>"
                    + "<div>new line ends the quote</div>"
                ),
                (
                    "prefix<br><blockquote>block quotes with <strong>strong <em>and italic</em> text</strong>"
                    + " work</blockquote><br>new line ends the quote<br>"
                )
            ],
            [
                (
                    "text <ts-text-command class=\"ts-is-valid\">"
                    + "https://example.com/?q=*not-strong*</ts-text-command>"
                    + " *strong* text"
                ),
                (
                    "text <ts-text-command class=\"ts-is-valid\">"
                    + "https://example.com/?q=*not-strong*</ts-text-command>"
                    + " <strong>strong</strong> text"
                )
            ],
            [
                (
                    "text <ts-text-command class=\"ts-is-valid\">"
                    + "https://example.com/?q=*not-strong*</ts-text-command>"
                    + " *strong* text"
                ),
                (
                    "text <ts-text-command class=\"ts-is-valid\">"
                    + "https://example.com/?q=*not-strong*</ts-text-command>"
                    + " <strong>strong</strong> text"
                )
            ],
            [
                (
                    "text <a href=\"https://example.com/?q=*not-strong*\"\n"
                    + "  title=\"https://example.com/?q=*not-strong*\">"
                    + "https://example.com/?q=*not-strong*</a>"
                    + " *strong* text"
                ),
                (
                    "text <a href=\"https://example.com/?q=*not-strong*\""
                    + " title=\"https://example.com/?q=*not-strong*\">"
                    + "https://example.com/?q=*not-strong*</a>"
                    + " <strong>strong</strong> text"
                )
            ],
            [
                "<div class=\"x\">malformed *html</div><div> *strong* text",
                "<div class=\"x\">malformed *html</div><br> <strong>strong</strong> text<br>"
            ],
            [
                "html ```force-closes <p> code ``` block",
                "html <br><pre><code>force-closes </code></pre><br><p> code ``` block</p>"
            ],
            [
                "mixed ```markup <a href=\"#\">``` and HTML</a> text",
                "mixed <br><pre><code>markup </code></pre><br><a href=\"#\">``` and HTML</a> text"
            ],
            [
                "mixed ```markup <a href=\"#\">``` and HTML</a> text",
                "mixed <br><pre><code>markup </code></pre><br><a href=\"#\">``` and HTML</a> text"
            ],
            [
                (
                    "<div itemprop=\"copy-paste-block\" class=\"copy-paste-block\">"
                    + "<div>line 1<br>*line 2*</div><div>line 3<br>"
                    + "&gt; quote<br>"
                    + "```&nbsp;&nbsp;&nbsp; code line 1<br>"
                    + "&nbsp;&nbsp;&nbsp; code line 2```"
                    + "<br>line 4<br></div></div>"
                ),
                (
                    "<div itemprop=\"copy-paste-block\" class=\"copy-paste-block\">"
                    + "line 1<br><strong>line 2</strong><br>line 3<br>"
                    + "<blockquote>quote</blockquote><br>"
                    + "<br><pre><code>&nbsp;&nbsp;&nbsp; code line 1<br>"
                    + "&nbsp;&nbsp;&nbsp; code line 2</code></pre><br>"
                    + "<br>line 4<br></div>"
                )
            ],
            [
                "text<br>&gt; quote ```code block``` text",
                "text<br><blockquote>quote </blockquote><br><pre><code>code block</code></pre><br> text"
            ],
            [
                "text &gt; not quote",
                "text &gt; not quote"
            ],
            [
                "custom emoji: :pm: <a href=\"https://example.com/?a=:pm:\">:pm:</a> `:pm:`",
                (
                    "custom emoji: "
                    + "<img src=\"" + custom_emojis["pm"]["src"] + "\""
                    + " alt=\":pm:\" title=\":pm:\""
                    + " width=\"" + custom_emojis["pm"]["width"] + "\""
                    + " height=\"" + custom_emojis["pm"]["height"] + "\" />"
                    + " <a href=\"https://example.com/?a=:pm:\">:pm:</a> <code>:pm:</code>"
                )
            ],
            [
                (
                    "<div itemprop=\"copy-paste-block\" class=\"copy-paste-block\">"
                    + "<div>text<br>text<div>text<br type=\"moz\">text</div></div>"
                    + "</div>"
                ),
                (
                    "<div itemprop=\"copy-paste-block\" class=\"copy-paste-block\">"
                    + "text<br>text<br>text<br type=\"moz\">text<br>"
                    + "</div>"
                )
            ],
            [
                (
                    "<div>line 1<br></div>"
                    + "<div>line 2<br></div>"
                    + "<div>```code 1<br></div>"
                    + "<div>code 2<br></div>"
                    + "<div>code 3```<br></div>"
                    + "<div>line 3<br></div>"
                ),
                (
                    "line 1<br>"
                    + "line 2<br>"
                    + "<br>"
                    + "<pre><code>code 1<br>"
                    + "code 2<br>"
                    + "code 3</code></pre><br>"
                    + "<br>"
                    + "line 3<br>"
                )
            ],
            [
                "<div>text <ts-text-command class=\"\">@Mentio</ts-text-command><br></div>",
                "text <ts-text-command class=\"\">@Mentio</ts-text-command><br>"
            ],
            [
                (
                    "<div>text <span itemscope=\"\" itemtype=\"http://schema.skype.com/Mention\""
                    + " data-itemprops=\"{&quot;mri&quot;:&quot;11:11111111111111111111111111111111"
                    + "@thread.tacv2&quot;,&quot;mentionType&quot;:&quot;team&quot;,&quot;memberCount"
                    + "&quot;:0}\">Mention</span> text<br></div>"
                ),
                (
                    "text <span itemscope=\"\" itemtype=\"http://schema.skype.com/Mention\""
                    + " data-itemprops=\"{&quot;mri&quot;:&quot;11:11111111111111111111111111111111"
                    + "@thread.tacv2&quot;,&quot;mentionType&quot;:&quot;team&quot;,&quot;memberCount"
                    + "&quot;:0}\">Mention</span> text<br>"
                )
            ],
            [
                "<div class=\"foo\" ng-click=\"whatever()\"><!-- comment --> text</div>",
                "<div class=\"foo\" ng-click=\"whatever()\"><!-- comment --> text</div>"
            ],
            [
                (
                    "<card-enlarge-image-icon><!----><a ng-show=\":wtf:\">"
                    + "<svg viewBox=\"0 0 32 32\" role=\"presentation\""
                    + " class=\"app-svg icons-expand\" focusable=\"false\">"
                    + "<g class=\"icons-default-fill\"><path class=\"icons-unfilled\" d=\"1,2,\n3,4\">"
                    + "</path></g></svg>:wtf: *should remain*</a></card-enlarge-image-icon>"
                ),
                (
                    "<card-enlarge-image-icon><!----><a ng-show=\":wtf:\">"
                    + "<svg viewBox=\"0 0 32 32\" role=\"presentation\""
                    + " class=\"app-svg icons-expand\" focusable=\"false\">"
                    + "<g class=\"icons-default-fill\"><path class=\"icons-unfilled\" d=\"1,2,\n3,4\">"
                    + "</path></g></svg>:wtf: *should remain*</a></card-enlarge-image-icon>"
                )
            ],
            [
                "<div>text 1 ```code 1<br></div><div>code 2<br></div><div>code 3```<br></div><div>text 2<br></div>",
                "text 1 <br><pre><code>code 1<br>code 2<br>code 3</code></pre><br><br>text 2<br>"
            ],
            [
                "<!-- *comment* --> *markdown*",
                "<!-- *comment* --> <strong>markdown</strong>",
            ],
            [
                "text ```code <a href=\"#\">``` this isn't code</a>",
                "text <br><pre><code>code </code></pre><br><a href=\"#\">``` this isn't code</a>",
            ],
            [
                (
                    "<div>text ```<br></div>"
                    + "<div itemprop=\"copy-paste-block\" class=\"copy-paste-block\">"
                    + "<div>function test()<br>"
                    + "{\u200b<br>"
                    + "&nbsp;&nbsp;&nbsp; var foo;</div>"
                    + "<div>&nbsp;&nbsp;&nbsp; bar = [];</div>"
                    + "<div>&nbsp;&nbsp;&nbsp; for (foo in baz) {\u200b<br>"
                    + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; if (baz.hasOwnProperty(foo)) {\u200b<br>"
                    + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                    + " bar.push(foo.substr(0, 1));<br>"
                    + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; }\u200b<br>"
                    + "&nbsp;&nbsp;&nbsp; }\u200b</div>"
                    + "<div>&nbsp;&nbsp;&nbsp; bar = bar.join(\"\");<br>"
                    + "}\u200b```<br></div>"
                    + "<div>asdasd<br></div>"
                    + "</div>"
                ),
                (
                    "text <br>"
                    + "<pre><code><br>"
                    + "<div itemprop=\"copy-paste-block\" class=\"copy-paste-block\">function test()<br>"
                    + "{\u200b<br>"
                    + "&nbsp;&nbsp;&nbsp; var foo;<br>"
                    + "&nbsp;&nbsp;&nbsp; bar = [];<br>"
                    + "&nbsp;&nbsp;&nbsp; for (foo in baz) {\u200b<br>"
                    + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; if (baz.hasOwnProperty(foo)) {\u200b<br>"
                    + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                    + " bar.push(foo.substr(0, 1));<br>"
                    + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; }\u200b<br>"
                    + "&nbsp;&nbsp;&nbsp; }\u200b<br>"
                    + "&nbsp;&nbsp;&nbsp; bar = bar.join(\"\");<br>"
                    + "}\u200b</code></pre><br>"
                    + "<br>"
                    + "asdasd<br></div>"
                )
            ]
        ];

        QUnit.module("md_to_html", function () {
            QUnit.test("md_to_html", function(assert) {
                var i, l, input, expected;

                for (i = 0, l = tests.length; i < l; ++i) {
                    input = tests[i][0];
                    expected = tests[i][1];

                    assert.equal(
                        md_to_html(input),
                        expected,
                        "input: " + JSON.stringify(input)
                    );
                }

            });
        });
    }

    function handle_key(editor, evt_name, evt)
    {
        var dom_evt, editor_element, key, code, emoji_dropdowns;

        if (!is_enabled) {
            return;
        }

        log_debug("handle_key, evt_name=" + evt_name);

        dom_evt = evt.data.domEvent.$;

        if (typeof(dom_evt) === "object") {
            key = String(dom_evt.key);
            code = String(dom_evt.code).toLowerCase();
        } else {
            log_debug("cannot find domEvent, skipping");
            return;
        }

        log_trace("dom_evt.key = " + key + " dom_evt.code = " + code);

        if (code === "enter") {
            if (dom_evt.shiftKey) {
                log_debug("Shift+Enter pressed, still editing");
            } else {
                log_debug("Enter pressed");

                editor_element = editor.document.$.activeElement;

                if (editor_element.innerHTML.match(/<ts-text-command[^>]*>/)) {
                    log_debug(
                        "incomplete text-command found, ignoring Enter key"
                        + " in order to avoid interfering with the autocomplete"
                    );
                } else {
                    emoji_dropdowns = document.getElementsByClassName(
                        "dropdown-menu ts-emoji-dropdown-wrapper"
                    );

                    if (emoji_dropdowns.length > 0) {
                        log_debug(
                            "emoji dropdown might be active, not doing the formatting yet"
                            + " in order to avoid interfering with the autocomplete"
                        );
                    } else {
                        log_debug("formatting message before sending");
                        editor_element.innerHTML = md_to_html(editor_element.innerHTML);
                    }
                }
            }
        } else if (
            previous_key === ":"
            && key.length === 1
            && custom_emoji_initials.indexOf(key) > -1
        ) {
            log_debug("preventing emoji which might collide with custom emojis")
            evt.stop();
            stop_change_events = 3;
        } else if (key === ">") {
            log_debug("preventing blockquote wysiwyg")
            evt.stop();
            stop_change_events = 3;
        } else if (md_chars.indexOf(key) > -1) {
            log_debug("markdown key pressed, stopping the next change event");
            stop_change_events = 3;
        }

        if (key.length === 1) {
            previous_key = key;
        }
    };

    function handle_change(editor, evt_name, evt)
    {
        if (!is_enabled) {
            return;
        }

        log_debug("handle_change, evt_name=" + evt_name);
        log_trace("innerHTML: " + JSON.stringify(editor.document.$.activeElement.innerHTML));

        if (stop_change_events > 0) {
            log_debug("skipping change event listeners");
            --stop_change_events;
            evt.stop();
        }
    };

    function add_handler(editor, evt_name, fn)
    {
        editor.on(
            evt_name,
            function (evt) { return fn(editor, evt_name, evt); },
            undefined,
            undefined,
            -9999
        );
    }

    function patch_editor(editor)
    {
        log_debug("adding event handlers");

        add_handler(editor, "key", handle_key);
        add_handler(editor, "keyup", handle_key);
        add_handler(editor, "keypress", handle_key);
        add_handler(editor, "change", handle_change);
    }

    function patch_editors()
    {
        var instances = _window.CKEDITOR.instances,
            key;

        stop_change_events = 0;

        _window.CKEDITOR.on(
            "instanceReady",
            function (evt) { return patch_editor(evt.editor); },
            undefined,
            undefined,
            -9999
        );

        for (key in instances) {
            if (instances.hasOwnProperty(key)) {
                log_debug("patching existing editor instance: " + String(key));
                patch_editor(instances[key]);
            }
        }
    }

    function try_patching_editors()
    {
        if (typeof(_window.CKEDITOR) === "object") {
            log_info("editor found, patching");
            patch_editors();
        } else {
            ++failures;

            if (failures < 30) {
                log_info("editor not found, retrying soon");
                setTimeout(try_patching_editors, 1000);
            } else {
                log_info("editor not found, too many failures, giving up");
            }
        }
    }

    function initialize_custom_emojis()
    {
        var name;

        custom_emoji_initials = [];

        for (name in custom_emojis) {
            if (custom_emojis.hasOwnProperty(name)) {
                custom_emoji_initials.push(name.substr(0, 1));
            }
        }

        custom_emoji_initials = custom_emoji_initials.join("");
    }

    function initialize_settings()
    {
        function reset_hiding_timeout()
        {
            if (hiding_timeout !== null) {
                clearTimeout(hiding_timeout);
                hiding_timeout = null;
            }
        }

        function stop_hiding()
        {
            div.style.bottom = "5px";
            reset_hiding_timeout();
        }

        var div = document.createElement("div"),
            a = document.createElement("a"),
            hiding_timeout = null;

        div.style.position = "fixed";
        div.style.zIndex = "10000";
        div.style.opacity = "1";
        div.style.right = "5px";
        div.style.bottom = "5px";
        div.style.margin = "0";
        div.style.padding = "5px";
        div.style.border = "outset 2px #6264a7";
        div.style.borderRadius = "5px";
        div.style.background = "#6264a7";
        div.style.display = "block";
        div.style.cursor = "pointer";
        div.style.width = "auto";
        div.style.height = "auto";
        div.style.minHeight = "20px";
        div.style.transition = "bottom 0.3s ease-in-out";

        div.addEventListener(
            "mouseover",
            function ()
            {
                hiding_timeout = setTimeout(
                    function ()
                    {
                        if (hiding_timeout !== null) {
                            div.style.bottom = "-28px";
                            hiding_timeout = null;
                        }
                    },
                    1600
                );
            }
        );
        div.addEventListener(
            "mouseleave",
            function ()
            {
                reset_hiding_timeout();
                setTimeout(stop_hiding, 4100);
            }
        );

        a.style.position = "auto";
        a.style.zIndex = "10000";
        a.style.opacity = "1";
        a.style.margin = "0";
        a.style.padding = "0";
        a.style.border = "none";
        a.style.background = "#6264a7";
        a.style.color = "#fff";
        a.style.fontFamily = "sans-serif";
        a.style.fontSize = "12px";
        a.style.fontWeight = "normal";
        a.style.lineHeight = "20px";
        a.style.height = "20px";
        a.style.display = "inlnie";
        a.style.textDecoration = "none";

        a.innerHTML = "wysimd ON";

        a.addEventListener(
            "click",
            function (evt) {
                var on_off;

                is_enabled = !is_enabled;
                on_off = is_enabled ? "ON" : "OFF";
                a.innerHTML = "wysimd " + on_off;
                log_info("turned " + on_off);

                evt.stopPropagation();
                evt.preventDefault();
                stop_hiding();
            }
        );

        div.appendChild(a);
        document.getElementsByTagName("body")[0].appendChild(div);
    }

    function turn_off_wysiwyg()
    {
        initialize_custom_emojis();
        initialize_settings();
        try_patching_editors();
    }

    function run_tests()
    {
        test_md_to_html();
    }

    function main()
    {
        if (window.location.hostname.match(/^teams\.microsoft\.com\.?$/)) {
            log_info("trying to turn off WYSIWYG editors");
            turn_off_wysiwyg();
        } else if (typeof(QUnit) === "object") {
            log_info("running tests");
            run_tests();
        } else {
            log_info("how did we end up here?");
        }
    }

    main();
}

/**
 * The editor needs to have permission to call our event handlers.
 */
window.eval("(" + String(WYSIMD) + ")();");

})()
