export function openFeedback (context) {
// var report = function(context) {
  openUrl('https://github.com')
}

function openUrl (url) {
  NSWorkspace.sharedWorkspace().openURL(NSURL.URLWithString(url))
}
