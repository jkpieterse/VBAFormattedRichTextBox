<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="en">

<!-- #BeginTemplate "../jkp-ads_aspx.dwt" -->

<head>
<meta content="width=device-width, initial-scale=1" name="viewport" />
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<!-- #BeginEditable "doctitle" -->
<title>All-VBA Rich Text Box control for Excel, Word and Access</title>
<!-- #EndEditable -->
<!-- #BeginEditable "headsection" -->
<meta content="An all-VBA Rich Text Box look-alike for MSForms UserForms, with HTML-like markup, per-character formatting, text and markup editing, keyboard shortcuts and no external RichEdit control." name="description" />
<meta content="VBA Rich Text Box, formatted VBA text box, Excel Rich Text Box, Word Rich Text Box, Access Rich Text Box, MSForms rich text, HTML-like markup, character formatting, all VBA control, 64 bit Office" name="keywords" />
<meta content="en-us" http-equiv="Content-Language" />
<!-- #EndEditable --><% Response.Write(SiteTools.GetArticlejson()) %>
<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "Organization",
  "name": "JKP Application Development Services",
  "url": "https://jkp-ads.com/",
  "logo": "https://jkp-ads.com/images/jkp-adslogo.gif",
  "founder": {
    "@type": "Person",
    "name": "Jan Karel Pieterse",
    "url": "https://jkp-ads.com/about.aspx"
  },
  "contactPoint": {
    "@type": "ContactPoint",
    "email": "info@jkp-ads.com",
    "contactType": "customer support",
    "availableLanguage": ["en", "nl"]
  },
  "sameAs": ["https://www.linkedin.com/in/jankarelpieterse"]
}
</script>
<!-- #BeginEditable "scripting" -->
<!-- No article-specific scripting is required. The control itself is implemented in VBA. -->
<!-- #EndEditable -->
<script src="/includes/topnav_script.js" type="text/javascript"></script>
<meta content="General" name="rating" />
<meta content="no" http-equiv="imagetoolbar" />
<meta content="Copyright ©, Jan Karel Pieterse All Rights Reserved" name="copyright" />
<meta content="nocache" name="robots" />
<meta content="noarchive" name="robots" />
<link href="../jkp-ads.css" rel="stylesheet" type="text/css" />
<% Response.Write(SiteTools.GetCanonical()) %>
</head>

<body id="top">
<div id="container">
	<div id="banner">
		<div id="logoleft">
			<a href="../index.aspx"><img alt="Home" src="../images/jkp-adslogo.gif" /></a>
		</div>
		<div id="logomiddle">
			<div class="search-container" role="search">
				<form action="https://www.google.com/search" method="get">
					<input name="ie" type="hidden" value="UTF-8" />
					<input name="oe" type="hidden" value="UTF-8" />
					<div class="smtxt"><label for="googlesearch">Site search</label></div>
					<input id="googlesearch" maxlength="255" name="q" placeholder="Search.." size="12" type="text" value="" />
					<button class="dropbtn" name="btnG" type="submit" value="?">?</button>
					<input name="domains" type="hidden" value="https://jkp-ads.com" />
					<input hidden="true" name="sitesearch" type="radio" value="" />
					<input checked="checked" hidden="true" name="sitesearch" type="radio" value="https://jkp-ads.com" />
				</form>
			</div>
		</div>
	</div>
	<% Response.write(SiteTools.GetNav()) %><br style="clear: both" />
	<% Response.write(SiteTools.GetCrumbs()) %><br />
	<div id="content" role="main">
		<%
If FeatureToggle.IsEnabled("AnnouncementEnabled") Then
    Dim lang = FeatureToggle.GetCurrentLanguage(HttpContext.Current)
    Dim heading = FeatureToggle.GetMessage("AnnouncementHeading", lang)
    Dim body = FeatureToggle.GetMessage("AnnouncementBody", lang)
%>
		<div style="border: 1px solid #ccc; background-color: #f9f9f9; padding: 15px; margin: 15px 0; border-left: 4px solid #d9534f;">
			<h2><%= HttpUtility.HtmlEncode(heading) %></h2>
			<p><%= HttpUtility.HtmlEncode(body) %></p>
		</div>
		<% End If %>
		<!-- #BeginEditable "content" -->
		<h1>An all-VBA Rich Text Box control for Excel, Word and Access</h1>
		<h2 id="Content0">Content</h2>
		<ul>
			<li><a href="#Introduction">Introduction</a></li>
			<li><a href="#Features">Rich Text Box features</a></li>
			<li><a href="#Files">Files to import</a></li>
			<li><a href="#Setup">Add the control to a UserForm</a></li>
			<li><a href="#HowItWorks">How the Rich Text Box works</a></li>
			<li><a href="#Markup">Markup syntax</a></li>
			<li><a href="#Formatting">Format characters in code</a></li>
			<li><a href="#Editing">Text and markup editing</a></li>
			<li><a href="#Selection">Keyboard control and focus</a></li>
			<li><a href="#Lifecycle">Resizing and cleanup</a></li>
			<li><a href="#Compatibility">Compatibility</a></li>
			<li><a href="#Disclaimer">Disclaimer</a></li>
			<li><a href="#faq">Frequently asked questions</a></li>
		</ul>
		<h2 id="Introduction">Introduction</h2>
		<p>The standard MSForms TextBox is useful for ordinary text, but it cannot display different formatting within the same value. A real RichEdit control would solve that problem, but it would also introduce an external ActiveX dependency, registration issues and platform differences.</p>
		<p>This all-VBA Rich Text Box look-alike takes a smaller route. It uses a standard MSForms Frame as its host and creates ordinary MSForms labels for the formatted display. A temporary TextBox is used when the user edits the text or the markup.</p>
		<p>The result is a reusable control with no RichEdit reference and no Windows API dependency. The source files are deliberately kept together so you can copy the control into your own UserForm project.</p>
		<h2 id="Features">Rich Text Box features</h2>
		<ul>
			<li>HTML-like markup for bold, italic, underline and strikethrough text.</li>
			<li>Font name, font size, foreground color and background color markup.</li>
			<li>One-based <code>Characters</code> ranges for programmatic formatting.</li>
			<li>Plain-text editing and source/markup editing.</li>
			<li>Formatting-preserving plain-text edits where unchanged text keeps its format.</li>
			<li>F2, Ctrl+Enter, Escape and markup-toggle keyboard shortcuts.</li>
			<li>A hidden focus proxy so a host UserForm can keep keyboard focus on the control.</li>
			<li>A visible host-frame focus border and explicit cleanup of generated controls.</li>
		</ul>
		<h2 id="Files">Files to import</h2>
		<p>The complete reusable control consists of two class modules:</p>
		<ul>
			<li><code>clsRichTextBox.cls</code>: The controller, parser, renderer and editor lifecycle.</li>
			<li><code>clsRichTextCharacters.cls</code>: The one-based range object used for code-driven formatting.</li>
		</ul>
		<h2 id="Setup">Add the control to a UserForm</h2>
		<p>The UserForm must contain a design-time MSForms Frame. Keep the mandatory host routines together: declare the control, assign <code>HostFrame</code>, forward the Frame's Enter, Exit and Resize events, focus the control on form activation, and call <code>Terminate</code> before the form is unloaded.</p>
		<pre><code>Option Explicit

Private WithEvents mcRichText As clsRichTextBox

Private Sub UserForm_Initialize()
    Set mcRichText = New clsRichTextBox
    Set mcRichText.HostFrame = Me.frRichText
    mcRichText.Markup = "This is &lt;b&gt;bold&lt;/b&gt; and &lt;i&gt;italic&lt;/i&gt;."
End Sub

Private Sub UserForm_Activate()
    If Not mcRichText Is Nothing Then mcRichText.SetFocus
End Sub

Private Sub frRichText_Enter()
    If Not mcRichText Is Nothing Then mcRichText.EnterExit False
End Sub

Private Sub frRichText_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not mcRichText Is Nothing Then mcRichText.EnterExit True
End Sub

Private Sub frRichText_Resize()
    If Not mcRichText Is Nothing Then mcRichText.Resize
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not mcRichText Is Nothing Then mcRichText.Terminate
    Set mcRichText = Nothing
End Sub</code></pre>
		<p>The Frame name in this example is <code>frRichText</code>. Use your own Frame name in all mandatory routines. The buttons and status label in the demo are optional host code, not part of the minimum control integration.</p>
		<h2 id="HowItWorks">How the Rich Text Box works</h2>
		<p>The control stores plain text separately from a format array with one entry per character. Each entry contains the font name, size, bold, italic, underline, strikethrough, weight, charset, foreground color and background color.</p>
		<p>For display, adjacent characters with the same format are coalesced into one run. Each run is drawn as an ordinary MSForms label. The label's font is applied before <code>AutoSize</code>, so the next run can be placed directly beside it without Windows API text measurement.</p>
		<p>The parser starts with the host Frame's font name, font size and colors as defaults. Bold, italic, underline and strikethrough are off unless markup or code turns them on.</p>

		<h2 id="Markup">Markup syntax</h2>
		<p>The source editor accepts a deliberately small HTML-like syntax. It is not an HTML renderer. The following forms are supported:</p>
		<ul>
			<li><code>&lt;b&gt;bold&lt;/b&gt;</code> and <code>&lt;strong&gt;strong&lt;/strong&gt;</code>.</li>
			<li><code>&lt;i&gt;italic&lt;/i&gt;</code> and <code>&lt;em&gt;emphasis&lt;/em&gt;</code>.</li>
			<li><code>&lt;u&gt;underlined&lt;/u&gt;</code> and <code>&lt;s&gt;struck&lt;/s&gt;</code>.</li>
			<li><code>&lt;size=12&gt;large&lt;/size&gt;</code> and <code>&lt;font=Calibri&gt;text&lt;/font&gt;</code>.</li>
			<li><code>&lt;font color=&quot;#FF0000&quot;&gt;red&lt;/font&gt;</code>.</li>
			<li><code>&lt;span style=&quot;color:#FF0000;background-color:#FFFF00;font-size:12&quot;&gt;text&lt;/span&gt;</code>.</li>
		</ul>
		<p>To display a tag literally, escape its angle brackets. For example, <code>&amp;lt;b&amp;gt;bold&amp;lt;/b&amp;gt;</code> displays as <code>&lt;b&gt;bold&lt;/b&gt;</code> instead of applying bold formatting. Ampersands, less-than signs and greater-than signs are escaped when the control serializes markup.</p>
		<pre><code>mcRichText.Markup = "This is &lt;b&gt;bold&lt;/b&gt;, &lt;i&gt;italic&lt;/i&gt;, and &lt;font color=""#FF0000""&gt;red&lt;/font&gt;."
Debug.Print mcRichText.Text
Debug.Print mcRichText.Markup</code></pre>

		<h2 id="Formatting">Format characters in code</h2>
		<p><code>Characters</code> uses a 1-based start position and length, like Excel's Characters object. It returns a temporary range object with formatting properties. A mixed property value returns <code>Null</code>.</p>
		<pre><code>Dim oCharacters As clsRichTextCharacters

Set oCharacters = mcRichText.Characters(1, 8)
oCharacters.Bold = True
oCharacters.ForeColor = RGB(192, 0, 0)

mcRichText.Characters(10, 4).Italic = True
mcRichText.Characters(15, 10).Underline = True</code></pre>
		<p>Calling <code>Text</code> replaces the text and starts with the default format. Plain-text editing is different: unchanged prefix and suffix characters keep their formatting, and newly inserted text inherits the surrounding format where possible.</p>

		<h2 id="Editing">Text and markup editing</h2>
		<p>The control has two editing views. Plain-text editing uses a temporary multiline MSForms TextBox and commits through the formatting-preserving edit path. Markup editing puts the canonical markup in the same temporary TextBox and parses it when committed.</p>
		<p>Call <code>BeginEdit</code> for text editing or <code>BeginSourceEdit</code> for markup editing. <code>ToggleEditMode</code> switches between the two during one edit session. Invalid markup raises <code>ParseError</code> and leaves the last valid content intact.</p>
		<pre><code>Private Sub cmdEdit_Click()
    mcRichText.BeginEdit
End Sub

Private Sub cmdMarkup_Click()
    mcRichText.BeginSourceEdit
End Sub

Private Sub cmdToggle_Click()
    mcRichText.ToggleEditMode
End Sub

Private Sub cmdCommit_Click()
    mcRichText.CommitEdit
End Sub

Private Sub cmdCancel_Click()
    mcRichText.CancelEdit
End Sub</code></pre>

		<h2 id="Selection">Keyboard control and focus</h2>
		<p>The visible host is a Frame, which cannot reliably receive keyboard events itself. The control therefore creates a small hidden TextBox as a focus proxy. The host UserForm should call <code>SetFocus</code> on activation and forward the Frame's Enter and Exit events to <code>EnterExit</code>.</p>
		<ul>
			<li><code>F2</code>: start plain-text editing, or toggle modes while editing.</li>
			<li><code>Ctrl+Enter</code>: commit the current edit.</li>
			<li><code>Escape</code>: cancel the current edit.</li>
			<li><code>Ctrl+?</code> or <code>Ctrl+&lt;</code>: toggle text and markup editing where the keyboard layout reports the chord.</li>
		</ul>
		<p>Because punctuation key codes vary between keyboard layouts, the public <code>ToggleEditMode</code> method remains the reliable fallback. The host Frame border is changed while the control owns focus and restored when focus leaves.</p>
		<h2 id="Lifecycle">Resizing and cleanup</h2>
		<p>Call <code>Resize</code> after changing the host Frame's dimensions. Call <code>Terminate</code> when the UserForm is destroyed so all generated labels, focus proxies and edit TextBoxes are removed.</p>
		<pre><code>Private Sub UserForm_Resize()
    If Not mcRichText Is Nothing Then mcRichText.Resize
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not mcRichText Is Nothing Then mcRichText.Terminate
    Set mcRichText = Nothing
End Sub</code></pre>
		<p>The host Frame remains responsible for its own scrollbars and layout. The control deliberately does not require mouse-wheel hooks or an external RichEdit component.</p>
		<h2 id="Compatibility">Compatibility</h2>
		<p>The control contains no Windows API declarations and no Excel object-model 
		dependency. It is designed for MSForms projects, including 32-bit and 64-bit 
		Office. It is up to you to check behavior in each Office host and Mac version 
		required by your own deployment before distribution.</p>
		<h2 id="Disclaimer">Disclaimer</h2>
		<p>You use this control at your own risk. JKP Application Development Services 
		accepts no liability for damages arising from its use. Test the control 
		thoroughly in every Office host and platform supported by your project.</p>
		<h2 id="faq">Frequently asked questions</h2>
		<div class="smtxt">
			<p><a href="#Introduction">Why use an all-VBA Rich Text Box?</a></p>
			<p><a href="#HowItWorks">How are separately formatted runs displayed?</a></p>
			<p><a href="#Files">Which class modules must I import?</a></p>
			<p><a href="#Setup">How do I attach the control to a UserForm?</a></p>
			<p><a href="#Markup">Which markup tags are supported?</a></p>
			<p><a href="#Editing">How do text and markup editing work?</a></p>
			<p><a href="#Selection">Which keyboard shortcuts and focus events are supported?</a></p>
			<p><a href="#Compatibility">Does the control use Windows API calls?</a></p>
		</div>
		<div id="comments">
			<hr />
			<h2>Comments</h2>
			<div commentsection="/includes/getcomments.aspx?page=articles/vba-rich-text-box-control.aspx">
				Loading comments...</div>
			<div commentform="/includes/comments.aspx">
			</div>
		</div>
		<p>&nbsp;</p>
		<!-- #EndEditable -->
	</div>
	<div id="footer" role="contentinfo"><% Response.write(SiteTools.GetFooter()) %></div>
</div>
</body>

<!-- #EndTemplate -->
</html>
