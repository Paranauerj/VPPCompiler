// Add a new member function "Item" to the array object
// This function accepts a numerical index which then can
// be used to get the array item value.

Array.prototype.Item = function(idx) {
     return this[idx];
}

function GetJSONDataJS(fn) {
  return new Function('return ' + fn)();
}
