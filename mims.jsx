var MiMs = (function ($) {

    //Config object
    var MiMsConfig = {
        CacheKey: "murphy_mims_container",
        MimsResultsPageUrl: "/pages/mimsresults.aspx",
        SectionInfo: [
          {
              Title: "Policies & Strategy",
              TermSetGUID: "373df278-4204-4c57-b9a6-c2acc9bcfe28",
              ManagedProperty: "RefinableString33",
              DefaultClass: "policiesItem"
          },
          {
              Title: "Murphy Core Processes",
              TermSetGUID: "a9a5bb43-38bc-404d-9c7a-211343ae44c1",
              ManagedProperty: "RefinableString34",
              DefaultClass: ""
          },
          {
              Title: "Function Documentation",
              TermSetGUID: "50f6b622-9d26-4a94-aefa-068bfe13deaa",
              ManagedProperty: "RefinableString37",
              DefaultClass: "functionItem"
          },
          {
              Title: "Sector Documentation",
              TermSetGUID: "30029816-79f4-4095-a7f2-5ec702000075",
              ManagedProperty: "RefinableString21",
              DefaultClass: "sectionCapabilityItem"
          }
          //To be uncommented when the capability section has to be added back in.
          /*,
          {
              Title: "Capability Documentation",
              TermSetGUID: "ba8a66f2-36fc-4e3d-98bd-ca1b0832a919",
              ManagedProperty: "RefinableString38",
              DefaultClass: "sectionCapabilityItem"
          }*/]
    };


    var MimsLinkContainer = React.createClass({
        fetchTerms: function () {
            getTerms().done(function (containerJSON) {
                this.setState({ data: containerJSON });
            }.bind(this));

        },
        getInitialState: function () {
            return { data: [] };
        },
        componentDidMount: function () {
            this.fetchTerms();
        },
        render: function () {
            var sectionNodes = this.state.data.map(function (mimssection) {
                return (
          <MimsLinkSection title={mimssection.title} data={mimssection.data} key={mimssection.id}>
          </MimsLinkSection>
        );
            });

            return (
        <div className="MurphyMIMSContainer">
          <div>
              {sectionNodes}
          </div>
        </div>
      );
        }
    });

    var MimsLinkSection = React.createClass({
        render: function () {
            return (
        <div className="MimsLinkSection">
          <div className="fresh-wp-title">{this.props.title}</div>
          <MimsLinkList data={this.props.data} title={this.props.title} />
        </div>
      );
        }
    });

    var MimsLink = React.createClass({

        showArrow: function () {
            if (this.props.sectionTitle == "Murphy Core Processes") {
                return (
                    <div className="mimsArrow"></div>
                );
            }
        },
        render: function () {
            //console.log('background image render: ' + this.props.bgImage);
            var styles = {
                backgroundColor: this.props.bgColor,
                backgroundImage: 'url("' + this.props.bgImage + '")',
                backgroundRepeat: 'no-repeat',
                backgroundPosition: '8px 8px'
            };

            return (

        <div className="mimsLinkParent">
            {this.showArrow()}
            <div className={this.props.defaultClass}>
              <a href={this.props.link} style={styles}>{this.props.text}</a>
            </div>
        </div>
      );
        }
    });

    var MimsLinkList = React.createClass({
        processLabel: function (start) {
            if (this.props.title == "Murphy Core Processes") {
                return (
                            <div className="mimsLinkParent">
                                {start ? '' : <div className="mimsArrow"></div> }
                                <div className={ start ? 'processItemStart' : 'processItemEnd' }><div>{ start ? 'Start' : 'End' }</div></div>
                            </div>
                        );
            }
        },
        render: function () {
            var finalClass = "";
            if (this.props.title == "Murphy Core Processes") {
                finalClass = "processItemButton";
            }

            var sectionTitle = this.props.title;

            var LinkNodes = this.props.data.map(function (mimslink) {
                return (
                  <MimsLink link={mimslink.link} bgColor={mimslink.custombgColor} bgImage={mimslink.custombgImage} sectionTitle={sectionTitle} text={mimslink.text} key={mimslink.id} defaultClass={(finalClass != "") ? finalClass : mimslink.defaultClass + " MimsLink" }>
                  </MimsLink>
                );
            });

            return (
            <div className="mimsLinksList fresh-panel">
              <div className="mimsLinksListContainer">
                  {this.processLabel(true)}
                  {LinkNodes}
                  {this.processLabel(false)}
                  <div className="ms-clear"></div>
              </div>
            </div>
      );
        }
    });

    var init = function () {
        ReactDOM.render(<MimsLinkContainer />, document.getElementById("murphy-MIMS-LandingPage"));
    }

    //Get terms either from the cache or make a fresh call to the termstore if terms are not found in cache.
    var getTerms = function () {

        var deferred = new jQuery.Deferred();

        var containerJSON = getTermsFromCache()

        if (containerJSON == null) {
            getTermsFromSharePoint().done(function (containerJSON) {
                deferred.resolve(containerJSON);
            }).fail(function (errorMessage) {
                console.log(errorMessage);
            })
        }
        else {
            deferred.resolve(containerJSON);
        }

        return deferred.promise();

    }

    var getTermsFromCache = function () {
        return CC.CORE.Cache.Get(MiMsConfig.CacheKey);
    }

    var getTermsFromSharePoint = function () {

        var deferred = jQuery.Deferred();

        SP.SOD.executeFunc("sp.js", "SP.ClientContext", function () {
            SP.SOD.registerSod("sp.taxonomy.js", SP.Utilities.Utility.getLayoutsPageUrl("sp.taxonomy.js"));
            SP.SOD.executeFunc("sp.taxonomy.js", "SP.Taxonomy.TaxonomySession", function () {

                //Current Context
                var context = SP.ClientContext.get_current();

                //Current Taxonomy Session
                var taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);

                //Name of the Term Store from which to get the Terms.
                var termStore = taxSession.getDefaultSiteCollectionTermStore();
                var mimsTermSets = [];

                for (var i = 0; i < MiMsConfig.SectionInfo.length; i++) {
                    var termSet = termStore.getTermSet(MiMsConfig.SectionInfo[i].TermSetGUID);

                    mimsTermSets[i] = termSet.getAllTerms();

                    context.load(mimsTermSets[i], "Include(Id,Name,LocalCustomProperties)");
                }

                context.executeQueryAsync(function () {

                    var containerJSON = [];

                    var resultsPageUrl = _spPageContextInfo.webServerRelativeUrl + MiMsConfig.MimsResultsPageUrl;

                    for (var i = 0; i < mimsTermSets.length; i++) {

                        var section = {};
                        section.title = MiMsConfig.SectionInfo[i].Title;
                        section.data = [];
                        section.id = i;

                        var termEnumerator = mimsTermSets[i].getEnumerator();

                        while (termEnumerator.moveNext()) {

                            var currentTerm = termEnumerator.get_current();

                            var termLink = {}
                            termLink.id = currentTerm.get_id().toString();
                            termLink.text = currentTerm.get_name();
                            termLink.link = CC.CORE.Utilities.GetPreRefinedSearchPageUrl(resultsPageUrl, "", MiMsConfig.SectionInfo[i].ManagedProperty, currentTerm.get_name());

                            var customClass = currentTerm.get_localCustomProperties()["customcssclass"];

                            //Since each sector needs to have a different colour, the colour is attached as a local custom property to the terms in the sectors termset.
                            if (typeof customClass == 'undefined') {
                                termLink.defaultClass = MiMsConfig.SectionInfo[i].DefaultClass;
                            }
                            else {
                                termLink.defaultClass = customClass;
                            }

                            var custombgColor = currentTerm.get_localCustomProperties()["backgroundcolor"];
                            if (typeof custombgColor != 'undefined') {
                                termLink.custombgColor = custombgColor;
                                //console.log('background color: ' + termLink.custombgColor);
                            }

                            var custombgImage = currentTerm.get_localCustomProperties()["backgroundimage"];
                            if (typeof custombgImage != 'undefined') {
                                termLink.custombgImage = _spPageContextInfo.siteServerRelativeUrl + custombgImage;
                                //console.log('background image: ' + termLink.custombgImage);
                            }

                            section.data.push(termLink);
                        }

                        containerJSON.push(section);
                    }

                    CC.CORE.Cache.Set(MiMsConfig.CacheKey, containerJSON, CC.CORE.Cache.Timeout.Default);

                    deferred.resolve(containerJSON);

                }, function (sender, args) {

                    deferred.reject(args.get_message());

                });

            });
        });

        return deferred.promise();
    }

    return {
        Init: init
    }

})(jQuery);