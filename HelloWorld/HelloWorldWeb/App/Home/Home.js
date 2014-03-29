

/// <reference path="../App.js" />

    //"use strict";
    // Checks for the DOM to load using the jQuery ready function.
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // After the DOM is loaded, app-specific code can run.
            // Add any initialization logic to this function.
            var alreadyHasASave = Office.context.document.settings.get('saveAvaible');
            if (alreadyHasASave  == true) {
                loadEverything();
            }
            setTimeout(update, 1000)
            
        });
    };
    
    var reqCount = 1;
    var refCount = 1;
    var keyCount = 1;
    var bookCount = 1;
    var strCount = 1;
    var linkCount = 0;
    var textColour;
    var globalTemplate;
    var structSubPointsCount = new Array();
    structSubPointsCount[0] = 0;
    var document1 = "";

    function loadEverything() {
        deleteFirstBox('ref');
        deleteFirstBox('key');
        deleteFirstBox('book');
        deleteFirstBox('str');
        
        makeBoxes('ref');
        makeBoxes('key');
        makeBoxes('book');
        makeBoxes('str');
        var linkCount = Office.context.document.settings.get('linkCount');
        for (var i = 0; i < linkCount; i++) {
            var event = new String();
            event.keyCode = 13;
            event.which = 13;
            document.getElementById('text0').value = Office.context.document.settings.get('link' + i);
            converter(event, 'text0', 'mylist');
        }
        document.getElementById('text0').value = '';
        //load word count stuff
        document.getElementById('target').value = Office.context.document.settings.get('WC');
        document.getElementById('margin').value = Office.context.document.settings.get('margin');
        setTemplate(Office.context.document.settings.get('globalTemplate'));
        var backButton = Office.context.document.settings.get('back');
        setVisibility('back', backButton);
        setVisibility('help', !backButton);
        document.getElementById("category").innerText = Office.context.document.settings.get('title');

    }
    function deleteFirstBox(id){
        switch (id) {
            case 'ref':
               
                    removeBox('ref0');
                
                break;
            case 'key':
                
                    removeBox('key0');
                
                break;
            case 'book':
               
                    removeBox('book0');
                
                break;
            case 'str':
                
                    removeBox('str0');
                
                break;
        }
    }
    function makeBoxes(id) {
        var count = Office.context.document.settings.get(id+'Count');
        for (var i = 0; i < count; i++) {
            addInsElement(id);
            document.getElementById(id + i).value = Office.context.document.settings.get(id + i);
            var checkboxStatus = Office.context.document.settings.get(id + i + 'checkbox');
            if (checkboxStatus == true) {
                document.getElementById(id + i + 'checkbox').checked = true;
            }
            else if (checkboxStatus == false) {
                document.getElementById(id + i + 'checkbox').checked = false;
            }
            else {
                document.getElementById(id + i).value = "Something went very wrong here...";
            }
            switch(id){
                case 'str':
                    var subCount = Office.context.document.settings.get('str' + i + 'Count');
                    for (var j = 0; j < subCount; j++) {
                        var event = new String();
                        event.keyCode = 13;
                        event.which = 13;
                        addSubPoint(event, id + i)

                        document.getElementById('str' + i + '_sub' + j).value = Office.context.document.settings.get('str' + i + '_sub' + j);
                        document.getElementById('str' + i + '_sub' + j + 'checkbox').checked = Office.context.document.settings.get('str' + i + '_sub' + j + 'checkbox');
                    }
                    break;
            }
        }
    }
    function saveEverything() {
        //Office.context.document.settings.set('reqCount', reqCount);
        //for (var i = 0; i < reqCount; i++) {
        //    Office.context.document.settings.set('req' + i, document.getElementById('req' + i).value);
        //}
        Office.context.document.settings.set('saveAvaible', true);
        Office.context.document.settings.set('refCount', refCount);
        for (var i = 0; i < refCount; i++) {
            Office.context.document.settings.set('ref' + i, document.getElementById('ref' + i).value);
            Office.context.document.settings.set('ref' + i + 'checkbox', document.getElementById('ref' + i + 'checkbox').checked);
        }
        Office.context.document.settings.set('keyCount', keyCount);
        for (var i = 0; i < keyCount; i++) {
            Office.context.document.settings.set('key' + i, document.getElementById('key' + i).value);

            Office.context.document.settings.set('key' + i + 'checkbox', document.getElementById('key' + i + 'checkbox').checked);
        }
        Office.context.document.settings.set('bookCount', bookCount);
        for (var i = 0; i < bookCount; i++) {
            Office.context.document.settings.set('book' + i, document.getElementById('book' + i).value);
            Office.context.document.settings.set('book' + i + 'checkbox', document.getElementById('book' + i + 'checkbox').checked);
        }
        Office.context.document.settings.set('strCount', strCount);
        for (var i = 0; i < strCount; i++) {
            Office.context.document.settings.set('str' + i, document.getElementById('str' + i).value);
            Office.context.document.settings.set('str' + i + 'checkbox', document.getElementById('str' + i + 'checkbox').checked);

            Office.context.document.settings.set('str' + i + 'Count', structSubPointsCount[i]);
            for (var j = 0; j < structSubPointsCount[i]; j++) {
                Office.context.document.settings.set('str' + i + '_sub' + j, document.getElementById('str' + i + '_sub' + j).value);
                Office.context.document.settings.set('str' + i + '_sub' + j + 'checkbox', document.getElementById('str' + i + '_sub' + j + 'checkbox').checked);
            }
        }
        
        Office.context.document.settings.set('back', document.getElementById('back').style.display);
        Office.context.document.settings.set('help', document.getElementById('help').style.display);
        Office.context.document.settings.set('title', document.getElementById("category").innerText);
        Office.context.document.settings.set('WC', document.getElementById('target').value);
        Office.context.document.settings.set('margin', document.getElementById('margin').value);


        Office.context.document.settings.saveAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                //Hopefully it'll go in here (means that the save worked)
            } else {
                //Save didn't work
            }
        });
    }


    

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
    function converter(event, textBoxID, listID) {
        if (event.keyCode == 13 || event.which == 13) {
            var text = document.getElementById(textBoxID).value;
            if (text.length > 0) {
                if (text.substring(0, 7) != "http://" && text.substring(0, 8) != "https://") {
                    text = "http://" + text;
                }
                Office.context.document.settings.set('link' + linkCount++, text);
                var element = document.createElement("li");
                var link = document.createElement("a");
                link.setAttribute("href", text);
                link.setAttribute("target", "_blank");
                if (text.length <= 40) {
                    link.innerText = (text);
                }
                else {
                    link.innerText = text.substring(0, 40) + "...";
                    link.setAttribute("title", text);
                }
               // link.setAttribute("onclick", "openURL(link"+(linkCount-1)+")");
                element.appendChild(link);
                var list = document.getElementById("mylist");
                list.appendChild(element);
                document.getElementById(textBoxID).value = "";
                Office.context.document.settings.set('linkCount', linkCount);
                Office.context.document.settings.saveAsync(function (asyncResult) {
                    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                        //Hopefully it'll go in here (means that the save worked)
                    } else {
                        //Save didn't work
                    }
                });
            }
        }
    }

    //function openURL(name) {
    //    if (name.id == 'undefined') {
    //        name = name.id;
    //    }
    //    var link = Office.context.document.settings.get(name);
    //    var shell = new ActiveXObject("WScript.Shell");
    //    shell.run(link);
    //}
    function addInsElement(catagory) {
        // create a new insert element
        // and give it some content
        var newInp = document.createElement("input");
        var reqId;
        var form;
        switch (catagory) {
            case 'req':
                reqId = catagory + reqCount++;
                newInp.setAttribute("placeholder", "Enter a requirement here");
                form = "Requirements"; //name of the division
                break;
            case 'ref':
                reqId = catagory + refCount++;
                newInp.setAttribute("placeholder", "Enter a reference here");
                form = "References";//name of the division
                break;
            case 'str':
                structSubPointsCount[strCount] = 0;
                reqId = catagory + strCount++;
                newInp.setAttribute("placeholder", "Point " + strCount);
                newInp.setAttribute("onkeypress", "return addSubPoint(event," + reqId + ");")
                form = "Structure";//name of the division
                break;
            case 'key':
                reqId = catagory + keyCount++;
                newInp.setAttribute("placeholder", "Enter a key word here");
                form = "KeyWords";//name of the division
                break;
            case 'book':
                reqId = catagory + bookCount++;
                newInp.setAttribute("placeholder", "Enter a book title here");
                form = "Books"; //name of the division
                break;

                
        }
        newInp.setAttribute("type", "text");
        newInp.setAttribute("id", reqId);
        newInp.setAttribute("size", "29");
        newInp.setAttribute("MaxLength", "50");

        //create and insert corresponding checkbox
        var newCB = document.createElement("input");
        var CBid = reqId + "checkbox";

        newCB.setAttribute("type", "checkbox");
        newCB.setAttribute("id", CBid);
        newCB.setAttribute("tabindex", "-1");
        //create and insert coorresponding delete button
        var newButton = document.createElement("input");
        var butid = reqId + "button";

        newButton.setAttribute("type", "deleteButton");
        newButton.setAttribute("value", "X");
        newButton.setAttribute("id", butid);
        newButton.setAttribute("onclick", "removeBox(" + reqId + ");");
        newButton.setAttribute("readonly");
        newButton.setAttribute("tabindex", "-1");
        // add the newly created element and its content into the DOM
       
        my_div = document.getElementById(form);
        my_div.appendChild(newInp);
        my_div.appendChild(newCB);
        my_div.appendChild(newButton);
        //makes div for subpoints
        switch (catagory) {
            case 'req':

                break;
            case 'ref':

                break;
            case 'str':
                var newDiv = document.createElement("div");
                newDiv.setAttribute("id", reqId+"div");
                newDiv.setAttribute("style", "display:block;");
                my_div.appendChild(newDiv);
        }
        //BR because I can't br with "br"
        var p = document.createElement("p");
        my_div.appendChild(p);

    }

    function addSubPoint(event, catagory) {
        if (event.keyCode == 13 || event.which == 13) {
            var input;
            if (typeof catagory.id === 'undefined') {
                input = catagory;
            }
            else {
                input = catagory.id;
            }

            var newInp = document.createElement("input");
            var num = parseInt(input.match(/\d+/));
            var reqId;
            switch (input.replace(/[0-9]/g, '')) {
                case 'str':// str0/sub0
                    reqId = input + "_sub" + structSubPointsCount[num]++;
                    newInp.setAttribute("placeholder", "Subpoint " + structSubPointsCount[num]);
                    form = "Structure";//name of the division
                    break;
            }
           
            newInp.setAttribute("type", "text");
            newInp.setAttribute("id", reqId);
            newInp.setAttribute("size", "25");
            newInp.setAttribute("MaxLength", "50");
            newInp.setAttribute("style", "margin-left: 30px; height: 20px; line-height: 70%;")
            newInp.setAttribute("onkeypress", "return addSubPoint(event," + input + ");")
            //create and insert corresponding checkbox
            var newCB = document.createElement("input");
            var CBid = reqId + "checkbox";

            newCB.setAttribute("type", "checkbox");
            newCB.setAttribute("id", CBid);
            newCB.setAttribute("tabindex", "-1");
            newCB.setAttribute("style", "height: 16px; width: 16px;");
            //create and insert coorresponding delete button
            var newButton = document.createElement("input");
            var butid = reqId + "button";

            newButton.setAttribute("type", "deleteButton");
            newButton.setAttribute("value", "X");
            newButton.setAttribute("id", butid);
            newButton.setAttribute("onclick", "return removeSubBox(" + reqId + ");");
            newButton.setAttribute("readonly");
            newButton.setAttribute("tabindex", "-1");
            newButton.setAttribute("style", "height: 18px; width: 18px;");
            // add the newly created element and its content into the DOM

            my_div = document.getElementById(input + "div");
            my_div.appendChild(newInp);
            my_div.appendChild(newCB);
            my_div.appendChild(newButton);
            //BR because I can't br with "br"
            var p = document.createElement("p");
          //  my_div.appendChild(p);


        }
    }

    function writeToPage(text) {
        document.getElementById("Output").innerText = text;
    }

    //loops through all references and searches the whole document for each
    function refSearch(type) {
        var current;
        var needle;
        var n;
        var count;
        switch (type) {
            case 'ref':
                count = refCount;
                break;
            case 'key':
                count = keyCount;
                break;
        }
        for (var i = 0; i < count; i++) {
            current = type + i;
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

    

    function addSliceData(data) {
        document1 = document1.concat(data)
    }
    
    function progress() {     
        var total = 0;
        var id;
        //checks the references checkboxes
        var refChecked = countCheckboxes('ref');
        var x = document.getElementById("progressBarRef");
        var boxesRef = parseFloat(refChecked) / parseFloat(refCount);
        x.setAttribute("value", boxesRef);

        //checks the structure checkboxes
        var structChecked = countCheckboxes('str');
        x = document.getElementById("progressBarStruct");
        var boxesStr = parseFloat(structChecked) / parseFloat(strCount);
        x.setAttribute("value", boxesStr);

        //checks the book checkboxes
        var bookChecked = countCheckboxes('book');
        x = document.getElementById("progressBarBooks");
        var boxesBook = parseFloat(bookChecked) / parseFloat(bookCount);
        x.setAttribute("value", boxesBook);

        //checks the key word checkboxes
        var keyChecked = countCheckboxes('key');
        x = document.getElementById("progressBarKeyWords");
        var boxesKey = parseFloat(keyChecked) / parseFloat(keyCount);
        x.setAttribute("value", boxesKey);

        //updates the wordcount progress bar based on target wordcount
        var x = document.getElementById("progressBarWordCount");
        var value;
        if (document.getElementById('target').value == 0) 
            value = 1;
        else 
            value = wordCount() / document.getElementById('target').value;
        var y;
        if (document.getElementById('margin').value == "")
            y = 100;
        else
            y = parseInt(document.getElementById('margin').value) + 100
        if (value * 100 > y )
            x.style.color = "red";
        else
            x.style.color = "#4f9ad9";
        if (value > 1)
            value = 1;
        x.setAttribute("value", value);

        //updates the total progress bar, weights word count as 50% and requirements/references as 50%
        var x = document.getElementById("progressBarTotal");
        var total;
        switch (globalTemplate) {
            case 'Academic':
                total = (boxesRef * 0.25) + (boxesStr * 0.25) + (value * 0.5);
                break;
            case 'Foreign Language':
                total = (boxesRef * 0.2) + (boxesStr * 0.2) + (boxesKey * 0.2) + (value * 0.4);
                break;
            case 'Science':
                total = (boxesRef * 0.5) + (boxesStr * 0.5);
                break;
            case 'Creative Writing':
                total = (boxesStr * 0.5) + (value * 0.5);
                break;
        }
        x.setAttribute("value", total); 
    }

    function countCheckboxes(type) {
        var total;
        var count = 0;
        switch (type) {
            case 'str':
                total = strCount;
                break;
            case 'ref':
                total = refCount;
                break;
            case 'key':
                total = keyCount;
                break;
            case 'book':
                total = bookCount;
                break;
        }
        for (var i = 0; i < total; i++) {
            switch (type) {
                case 'str':
                    if (structSubPointsCount[i] > 0) {
                        var allChecked = true;
                        for (var j = 0; j < structSubPointsCount[i]; j++) {
                            if (document.getElementById('str' + i + '_sub' + j + 'checkbox').checked == false) {
                                allChecked = false;
                                break;
                            }
                        }
                        if (allChecked) {
                            document.getElementById('str' + i + 'checkbox').checked = true;
                        }
                    }
                    break;
                case 'ref':

                    break;
                case 'key':

                    break;
                case 'book':

                    break;
            }
            var current = type + i + "checkbox";
            var box = document.getElementById(current)
            if (box.checked) {
                count++;
            }
        }
        
        return count;
    }

    //reads in the file as text, searches it for references, counts the words, updates the progress bars and then deletes the file again.
    function update() {
        getFileData();
        calculateMargin();
        refSearch('ref');
        refSearch('key');
        progress();
        document1 = "";
        saveEverything();
        setTimeout(update, 500);
    };
    

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
            if(buttonID != ''){
                b.setAttribute("value", "Show");
            
            }
        }
        else {
            e.style.display = 'block';
            if (buttonID != '') {
                b.setAttribute("value", "Hide");
             
            }
        }
       
    }
    
    function setVisibility(id, bool) {
        var e = document.getElementById(id);
        if (bool) {
            e.style.display = 'block';
           
        }
        else {
            e.style.display = 'none';
      
        }
    }
    
    //word count function which replaces a unnecessary text and then counts the spaces in between words
    function wordCount() {
        s = document1;
        //s = s.replace(/(^\s*)|(\s*$)/gi, "");     //old method
        //s = s.replace(/[ ]{2,}/gi, " ");
        //s = s.replace(/\n /, "\n");
        ////writeToPage(s.split(' ').length);
        //return s.split(' ').length;
        try {                                         //new method
            //the try catch is incase there is a big change in the document, this can bring up an error sometimes but the try catch fixes this
            var word = s.match(/\S+/g);
            //document.getElementById("req0").value = ""+ word && word.length || 0; //Used for testing
            return word && word.length || 0;
        }
        catch (err) {
            return 0;
        }
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
    
    function setTemplate(template) {
        globalTemplate = template;
        Office.context.document.settings.set('globalTemplate', template);
        
        setVisibility('menu', false);
        setVisibility('help', false);

        setVisibility('template', true);
        switch (template) {
            case 'Foreign Language':
                document.getElementById("category").innerText = "Foreign Language";
                setVisibility('Books', false);
                setVisibility('BooksHead', false);
                setVisibility('linkList', false);
                setVisibility('LinksHead', false);
                break;
            case 'Science':
                document.getElementById("category").innerText = "Science";
                setVisibility('Books', false);
                setVisibility('BooksHead', false);
                setVisibility('WordCount', false);
                setVisibility('WordCountHead', false);
                setVisibility('KeyWords', false);
                setVisibility('KeyWordsHead', false);

                break;
            case 'Creative Writing':
                document.getElementById("category").innerText = "Creative Writing";
                setVisibility('Books', false);
                setVisibility('BooksHead', false);
                setVisibility('KeyWords', false);
                setVisibility('KeyWordsHead', false);
                setVisibility('References', false);
                setVisibility('RefHead', false);
                setVisibility('linkList', false);
                setVisibility('LinksHead', false);
                break;
            case 'Academic':
                document.getElementById("category").innerText = "Academic";
                setVisibility('KeyWords', false);
                setVisibility('KeyWordsHead', false);
                break;
        }

        Office.context.document.settings.saveAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                //Hopefully it'll go in here (means that the save worked)
            } else {
                //Save didn't work
            }
        });
    }

    function resetTemplate() {
        document.getElementById("category").innerText = "Menu";
        setVisibility('template', false);
        setVisibility('menu', true);
        setVisibility('helpSection', false);
        setVisibility('Books', true);
        setVisibility('BooksHead', true);
        setVisibility('WordCount', true);
        setVisibility('WordCountHead', true);
        setVisibility('KeyWords', true);
        setVisibility('KeyWordsHead', true);
        setVisibility('References', true);
        setVisibility('RefHead', true);
        setVisibility('Structure', true);
        setVisibility('StructureHead', true);
        setVisibility('linkList', true);
        setVisibility('LinksHead', true);
    }
    function loadHelp() {
        document.getElementById("category").innerText = "Help";
        setVisibility('template', false);
        setVisibility('menu', false);
        setVisibility('help', false);
        setVisibility('helpSection', true);
        setVisibility('back', true);
    }

    function removeSubBox(id) {
        var input;
        if (typeof id.id === 'undefined') {
            input = id;
        }
        else {
            input = id.id;
        }
        var box = document.getElementById(input);
        var checkBox = document.getElementById(input + "checkbox");
        var button = document.getElementById(input + "button");
        var catagories = input.replace(/[0-9]/g, '').split("_");  //splits name into two parts and removes numbers  str0/sub1 -> {str,sub}
        var numbers = input.replace(/[A-Za-z$-]/g, "").split("_");
        
        box.parentElement.removeChild(box);
        checkBox.parentElement.removeChild(checkBox);
        button.parentElement.removeChild(button);

        for (var i = (parseInt(numbers[1]) + 1) ; i < structSubPointsCount[numbers[0]]; i++) {
            document.getElementById(catagories[0] + numbers[0] + "_" + catagories[1] + i).setAttribute("placeholder", "Subpoint " + i);
            document.getElementById(catagories[0] + numbers[0] + "_" + catagories[1] + i).id = (catagories[0] + numbers[0] + "_" + catagories[1] + (i - 1));
            document.getElementById(catagories[0] + numbers[0] + "_" + catagories[1] + i + "checkbox").id = (catagories[0] + numbers[0] + "_" + catagories[1] + (i - 1)) + "checkbox";
            document.getElementById(catagories[0] + numbers[0] + "_" + catagories[1] + i + "button").setAttribute("onclick", "removeSubBox(" + (catagories[0] + numbers[0] + "_" + catagories[1] + (i - 1)) + ");");
            document.getElementById(catagories[0] + numbers[0] + "_" + catagories[1] + i + "button").id = (catagories[0] + numbers[0] + "_" + catagories[1] + (i - 1)) + "button";
        }

        structSubPointsCount[numbers[0]] -= 1;


    }



    //deletes normal boxes (doesn't work for sub boxes)
    function removeBox(id) {
        var input;
        if (typeof id.id === 'undefined') {
            input = id;
        }
        else {
            input = id.id;
        }
        var box = document.getElementById(input);
        var checkBox = document.getElementById(input + "checkbox");
        var button = document.getElementById(input + "button");
        var num = parseInt(input.match(/\d+/));
        var catagory = input.replace(/[0-9]/g, '');
        switch (catagory) {
            case 'str':
                var div = document.getElementById(catagory + num + "div");
                div.parentElement.removeChild(div);
        }
        box.parentElement.removeChild(box);
        checkBox.parentElement.removeChild(checkBox);
        button.parentElement.removeChild(button);
        var count;
        switch (catagory) {
            case 'req':
                count = reqCount;
                break;
            case 'ref':
                count = refCount;
                break;
            case 'str':
                count = strCount;
                break
            case 'key':
                count = keyCount;
                break;
            case 'book':
                count = bookCount;
                break;
        }
        for (var i = (num + 1) ; i < count; i++) {
            switch (catagory) {
                case 'req':
                    //incase of anything unique
                    break;
                case 'ref':
                    //incase of anything unique
                    break;
                case 'str':
                    document.getElementById(catagory + i).setAttribute("placeholder", "Point " + i);
                    for (var j = 0; j < structSubPointsCount[i]; j++) {
                        document.getElementById(catagory + i + "_sub" + j).setAttribute("onkeypress", "return addSubPoint(event," + catagory + (i - 1) + ");");
                        document.getElementById(catagory + i + "_sub" + j).id = (catagory + (i - 1) + "_sub" + j);
                    }
                    document.getElementById(catagory + i).setAttribute("onkeypress", "addSubPoint(event," + catagory + (i - 1) + ");");
                    document.getElementById(catagory + i + "div").id = (catagory + (i - 1)) + "div";
                    break;
                case 'key':
                    break;
                case 'book':
                    break;
            }
            document.getElementById(catagory + i).id = (catagory+ (i-1));
            document.getElementById(catagory + i + "checkbox").id = (catagory + (i - 1)) + "checkbox";
            document.getElementById(catagory + i + "button").setAttribute("onclick", "removeBox(" + (catagory + (i - 1)) + ");");
            document.getElementById(catagory + i + "button").id = (catagory + (i - 1)) + "button";
           
        }
            
            switch (catagory) {
                case 'req':
                    reqCount -= 1;
                    break;
                case 'ref':
                    refCount -= 1;
                    break;
                case 'str':
                    for (var i = num; i < strCount; i++) {
                        structSubPointsCount[i] = structSubPointsCount[i + 1];
                    }
                    strCount -= 1;
                    break;
                case 'key':
                    keyCount--;
                    break;
                case 'book':
                    bookCount--;
                    break;
            }
    }

    

   



    
