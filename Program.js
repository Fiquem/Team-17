// The initialize function is required for all apps.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Add any initialization logic to this function.
    });
}

var reqCount = 1;

function InputReq() {
    var min = document.getElementById('min').value;
    var max = document.getElementById('max').value;
    var reqs = "";
    for(var i = 0; i < reqCount; i++)
    {
        var current = "req" + i;
        reqs += document.getElementById(current).value;
    }
    document.getElementById("Requirements").innerText = min + max + reqs;
}

function addInsElement () {
  // create a new insert element
  // and give it some content
  var newInp = document.createElement("input");
  var reqId = "req" + reqCount;
  reqCount++;

  newInp.setAttribute("type", "text");
  newInp.setAttribute("id", reqId);
  newInp.setAttribute("size", "40");
  newInp.setAttribute("placeholder", "Enter a requirement here");
  newInp.setAttribute("MaxLength", "50");

  // add the newly created element and its content into the DOM
  my_div = document.getElementById("RequirementForm");
  document.body.insertBefore(newInp, my_div);

  // BR because I can't br with "br"
  var p = document.createElement("p");
  document.body.insertBefore(p, my_div);
}