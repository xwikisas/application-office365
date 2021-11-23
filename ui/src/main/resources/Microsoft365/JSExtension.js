function showDoc(url, paneId)
{
  const t = document.getElementById(paneId);
  if (t.docShown) {
    t.innerHTML = "";
    t.docShown = false;
  } else {
    t.innerHTML = "<br><iframe src=" + url + " width='100%' height='1000'></iframe>";
    t.docShown = true;
  }
}

function htmlEscape(s) {
  let v = document.createElement("span");
  v.innerText = s;
  return v.innerHTML;
}

require(['jquery','xwiki-meta'], function (jquery, xwikiMeta) {
  window.jquery = jquery;

  // TODO: make it a JSX, support multiple displays (so... number boxes)
  function doSearchMicrosoft365(number)
  {
    let form = jquery("#microsoft365-searchBox-" + number);
    form[0].disable();

    let submitButton = jquery("#microsoft365-searchBox-" + number + "  input[type='submit']")[0];
    submitButton.style.background = 'url("/xwiki/resources/icons/xwiki/spinner.gif")';
    submitButton.style["background-size"]="cover";

    let searchText = form.find("input[name^='searchText']")[0].value;
    let url = new XWiki.Document(XWiki.Model.resolve('xwiki:Microsoft365.DocumentList', XWiki.EntityType.DOCUMENT))
                .getURL("get", "searchText=_the_Text_to_Search_&format=json&outputSyntax=plain")
                .replace("_the_Text_to_Search_", escape(searchText));
    jquery.getJSON(url, function (results) {
      let r = document.getElementById('searchResult-' + number);
      if (results.error) {
        let p1 = document.createElement("p"), p2 = document.createElement("p"),
          b = document.createElement("b");
        p1.appendChild(b);
        b.innerText = results.error;
        p2.innerText = results.errorMessage;
        r.appendChild(r);
      } else {
        let counter = 0, s = "";
        // only keep sites with no results
        let nonEmptyResults = [];
        for(let r of results.results) {
          console.log("site: " + r.siteName)
          if(r.items && r.items.length>0)
            nonEmptyResults.push(r);
        }
        if (nonEmptyResults.length === 0) {
          s+= "<li><b>${escapetool.javascript($l11n.render('microsoft365.search.noResults'))}</b></li>";
        } else {
          let matchesHint = document.createElement("p");
          matchesHint.innerText =
              "${escapetool.javascript($services.localization.render('microsoft365.search.matching',['_my_query_here_']))}"
                  .replace("_my_query_here_", searchText);
          s += matchesHint.outerHTML;
          s += "<ul>";
          for(let result of nonEmptyResults) {
            result.items.forEach((function (item) {
              console.log(item.name);
              const previewFieldId = 'previewpane-' + number + '-' + counter;
              const saveURL = xwikiMeta.page + "?" +
                  "writeObject=do" +
                  "&nb=" + number +
                  "&editLink=" + escape(item.viewUrl) +
                  "&embedLink=" + escape(item.embedUrl) +
                  "&id=" + escape(item.id) +
                  (item.si? "site=" + escape(item.si): "") +
                  "&version=" + escape(item.version) +
                  "&fileName=" + escape(item.name);
              s+= "<li><a href='" + saveURL + "'>" + htmlEscape(item.name) +
                  " (${services.localization.render('microsoft365.embed')}) </a>&nbsp;" +
                  '(<a href="#" onclick="showDoc(\'' + item.embedUrl + '\',\'' + previewFieldId +
                  '\'); return false;">$services.localization.render("microsoft365.preview")</a>)' +
                  '<span id="'+ previewFieldId + '">&nbsp;</span>';
              counter++;
            }));          }
          s+= "</ul>";
          r.innerHTML = s;
        }
      }
      form[0].enable();
      submitButton.style.background = "";
    });
  }

  jquery(document).ready(function () {
    if(window.ms365BoxNumbers) {
      window.ms365BoxNumbers.forEach(function(nb) {
        jquery("#microsoft365-searchBox-" + nb).submit(function (evt) {
          evt.preventDefault();
          doSearchMicrosoft365(nb);
          return false;
        });
      });
    }
  });
});

