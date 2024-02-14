// index

for (const key in _entry_point_) {
  if (
    typeof _entry_point_[key] === 'function' &&
    !key.startsWith('on') &&
    !key.startsWith('do')
  ) {
    globalThis[key] = _entry_point_[key]
  }
}

// トリガー用の関数は実際に記述されている必要がある
function onOpen(e) {
  _entry_point_.onOpen(e)
}

// run で実行される関数は実際に記述されている必要がある
function doPrompt(msg) {
  return _entry_point_.doPrompt(msg)
}
function doPromptMulti(imageUrl, msg) {
  return _entry_point_.doPromptMulti(imageUrl, msg)
}
