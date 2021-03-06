// The initialize function is required for all apps.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add any initialization logic to this function.
    });
}

var reqCount = 1;
var refCount = 1;

function InputReq() {
    var min = document.getElementById('min').value;
    var max = document.getElementById('max').value;
    var reqs = "";
    for (var i = 0; i < reqCount; i++) {
        var current = "req" + i;
        reqs += document.getElementById(current).value;
    }
    document.getElementById("Requirements").innerText = min + max + reqs;
}

function addInsElement(catagory) {
    // create a new insert element
    // and give it some content
    var newInp = document.createElement("input");
    var reqId = catagory + ((catagory == 'req')?reqCount++:refCount++);

    newInp.setAttribute("type", "text");
    newInp.setAttribute("id", reqId);
    newInp.setAttribute("size", "40");
    newInp.setAttribute("placeholder", "Enter a " + ((catagory == 'req')?"requirement":"reference") + " here");
    newInp.setAttribute("MaxLength", "50");

    //create and insert corresponding checkbox
    var newCB = document.createElement("input");
    var CBid = reqId + "checkbox";

    newCB.setAttribute("type", "checkbox");
    newCB.setAttribute("id", CBid);

    // add the newly created element and its content into the DOM
    var form = (catagory == 'req') ? "RequirementForm" : "ReferenceForm"
    my_div = document.getElementById(form);
    document.body.insertBefore(newInp, my_div);
    document.body.insertBefore(newCB, my_div);

    // BR because I can't br with "br"
    var p = document.createElement("p");
    document.body.insertBefore(p, my_div);
}

function writeToPage(text) {
    document.getElementById("Requirements").innerText = text;
}

//loops through all references and searches the whole document for each
function refSearch() {
    for (var i = 0; i < refCount; i++) {
        var current = "ref" + i;
        writeToPage(document.getElementById(current).value);
        var needle = document.getElementById(current).value
        getFileData(needle, current);
    }
}


// Get all of the content from a Word document in 1KB chunks of text.
function getFileData(needle, id) {
    Office.context.document.getFileAsync(
    Office.FileType.Text,
    {
        sliceSize: 1000
    },
    function (asyncResult) {
        if (asyncResult.status === 'succeeded') {
            var myFile = asyncResult.value,
              state = {
                  file: myFile,
                  counter: 0,
                  sliceCount: myFile.sliceCount
              };
            getSliceData(state, needle, id, false);
        }
    });
}
// Get a slice from the file, as specified by
// the counter contained in the state parameter.
function getSliceData(state, needle, id, alreadyFound) {
    state.file.getSliceAsync(
      state.counter,
      function (result) {
          var slice = result.value,
            data = slice.data;
          state.counter++;
          // if the needle is contatined in this slice, set the checkbox, otherwise unset it
          // Check to see if the final slice in the file has
          // been reached—if not, get the next slice;
          // if so, close the file.
          var n = data.search(needle);
          if (n == -1 && alreadyFound == false) {
              id += "checkbox";
              document.getElementById(id).checked = false;
          }
          else {
              alreadyFound = true;
              id += "checkbox";
              document.getElementById(id).checked = true;
          }
          if (state.counter < state.sliceCount) {
              getSliceData(state);
          }
          else {
              closeFile(state, alreadyFound);
          }
      });
}
// Close the file when done with it.
function closeFile(state) {
    state.file.closeAsync(
      function (results) {
          // Inform the user that the process is complete.
      });
}
