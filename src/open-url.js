export function openFeedback (context) {
  openUrl('https://github.com/wousser/SketchExcelContentSync')
}

function openUrl (url) {
  NSWorkspace.sharedWorkspace().openURL(NSURL.URLWithString(url))
}