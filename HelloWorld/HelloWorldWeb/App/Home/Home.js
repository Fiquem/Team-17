/// <reference path="../App.js" />

    //"use strict";
    // Checks for the DOM to load using the jQuery ready function.
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // After the DOM is loaded, app-specific code can run.
            // Add any initialization logic to this function.
            
            $('#newReqButton').click(addInsElement);
            $('#inputReqButton').click(InputReq);

        });
    };
    
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
        newInp.setAttribute("size", "35");
        newInp.setAttribute("placeholder", "Enter a " + ((catagory == 'req')?"requirement":"reference") + " here");
        newInp.setAttribute("MaxLength", "50");

        //create and insert corresponding checkbox
        var newCB = document.createElement("input");
        var CBid = reqId + "checkbox";

        newCB.setAttribute("type", "checkbox");
        newCB.setAttribute("id", CBid);

        // add the newly created element and its content into the DOM
        var form = (catagory == 'req') ? "Requirements" : "References";
        my_div = document.getElementById(form);
        my_div.appendChild(newInp);
        my_div.appendChild(newCB);

        //BR because I can't br with "br"
        var p = document.createElement("p");
        my_div.appendChild(p);
    }

    function writeToPage(text) {
        document.getElementById("Output").innerText = text;
    }

    //loops through all references and searches the whole document for each
    function refSearch() {
        var current;
        var needle;
        var n;
        for (var i = 0; i < refCount; i++) {
            current = "ref" + i;
            needle = document.getElementById(current).value;
            if (needle != "") {
                n = document1.search(needle);
                if (n == -1) {
                    id = current + "checkbox";
                    document.getElementById(id).checked = false;
                }
                else {
                    id = current + "checkbox";
                    document.getElementById(id).checked = true;
                }
            }
            else {
                id = current + "checkbox";
                document.getElementById(id).checked = false;
            }
        }
    }


    // Get all of the content from a Word document in 1KB chunks of text.
    function getFileData() {
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
                getSliceData(state);
            }
        });
    }

    var progressCount = 0;
    // Get a slice from the file, as specified by
    // the counter contained in the state parameter.
    function getSliceData(state) {
        state.file.getSliceAsync(
          state.counter,
          function (result) {
              var slice = result.value,
                data = slice.data;
              state.counter++;
              //sends the data from the slice as text to the addSliceData function which concatenates it to the global variable document1
              addSliceData(data);

              // Check to see if the final slice in the file has
              // been reached—if not, get the next slice;
              // if so, close the file.
              if (state.counter < state.sliceCount) {
                  getSliceData(state);
              }
              else {
                  closeFile(state);
              }
          });
    }

    var document1 = "";

    function addSliceData(data) {
        document1 = document1.concat(data)
    }
    
    function progress() {     
        var x = document.getElementById("progressBarRef");
        var total = 0;
        var count = 0;
        var id;
        //checks the references checkboxes
        for (var i = 0; i < refCount; i++) {
            var current = "ref" + i + "checkbox";
            var box = document.getElementById(current)
            if (box.checked) {
               count++;
            }
        }
        var refChecked = count;
        x.setAttribute("value", parseFloat(count) / parseFloat(refCount));
        count = 0;

        //checks the requirements checkboxes
        for (var i = 0; i < reqCount; i++) {
            var current = "req" + i + "checkbox";
            var box = document.getElementById(current)
            if (box.checked) {
                count++;
            }
        }
        var reqChecked = count;
        var x = document.getElementById("progressBarReq");
        x.setAttribute("value", parseFloat(count) / parseFloat(reqCount));

        //updates the wordcount progress bar based on target wordcount
        var x = document.getElementById("progressBarWordCount");
        var value;
        if (document.getElementById('target').value == 0) 
            value = 1;
        else 
            value = wordCount() / document.getElementById('target').value;
        x.setAttribute("value", value);

        //updates the total progress bar, weights word count as 50% and requirements/references as 50%
        var x = document.getElementById("progressBarTotal");
        var boxesRef = (parseFloat(refChecked) / parseFloat(refCount));
        var boxesReq = (parseFloat(reqChecked) / parseFloat(reqCount));
        var total = (boxesRef * 0.25) + (boxesReq * 0.25) + (value * 0.5);
        x.setAttribute("value", total);
        
    }

    //reads in the file as text, searches it for references, counts the words, updates the progress bars and then deletes the file again.
    function update() {
        getFileData();
        calculateMargin();
        refSearch();
        progress();
        document1 = "";
    }


    // Close the file when done with it.
    function closeFile(state) {
        state.file.closeAsync(
          function (results) {
              // Inform the user that the process is complete.
          });
    }
    
    //used to show/hide divs
    function toggleVisibility(id, buttonID) {
        var e = document.getElementById(id);
        var b = document.getElementById(buttonID);
        if (e.style.display == 'block') {
            e.style.display = 'none';
            b.setAttribute("value", "Show");
        }
        else {
            e.style.display = 'block';
            b.setAttribute("value", "Hide");
        }
       
    }
    
    function setVisibility(id, bool) {
        var e = document.getElementById(id);
        if (bool)
            e.style.display = 'block';
        else
            e.style.display = 'none';
    }
    
    //word count function which replaces a unnecessary text and then counts the spaces in between words
    function wordCount() {
        s = document1;
        s = s.replace(/(^\s*)|(\s*$)/gi, "");
        s = s.replace(/[ ]{2,}/gi, " ");
        s = s.replace(/\n /, "\n");
        writeToPage(s.split(' ').length);
        return s.split(' ').length;
    }

    //simple function which calculates the upper and lower bounds of a word count given a target and percentage margin
    function calculateMargin() {
        var target = document.getElementById("target").value;
        if (target != 0) {
            var margin = document.getElementById("margin").value;
            var low = document.getElementById("low");
            var high = document.getElementById("high");

            var difference = target * (margin / 100);
            var lowerBound = target - difference;
            var upperBound = parseInt(target) + parseInt(difference);
            low.setAttribute("value", lowerBound);
            high.setAttribute("value", upperBound);
        }
    }

    
