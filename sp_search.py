import csv
import json
import requests
import sys



url = "https://substrate.office.com/search/api/v2/query"
bearer_token = ""  #//CHANGE ME//// a valid Bearer token is required

number_of_results = 50
sub_domain = "subdomain" #//CHANGE ME////   this is the subdomain for https://<subdomain>.sharepoint.com

headers = {"Authorization" : "Bearer " + bearer_token, "Content-Type" : "application/json"}


def search(searchterm,number_of_results):

    query = "{'EntityRequests':[{'EntityType':'File','ContentSources':['OneDriveBusiness','SharePoint'],'Fields':['.callerStack','.correlationId','.mediaBaseUrl','.spResourceUrl','.thumbnailUrl','AuthorOWSUSER','ClassicAttachmentVisualizationUrl','ContentClass','ContentTypeId','Created','DefaultEncodingURL','DepartmentId','Description','DocId','EditorOWSUSER','FileExtension','FileType','Filename','GeoLocationSource','HitHighlightedSummary','IsContainer','IsHubSite','LastModifiedTime','LinkingUrl','ListID','ListTemplateTypeId','MediaDuration','ModifiedBy','ParentLink','Path','PiSearchResultId','PictureThumbnailURL','ProgID','PromotedState','RelatedHubSites','SPSiteURL','SPWebUrl','SecondaryFileExtension','ServerRedirectedPreviewURL','ServerRedirectedUrl','SiteId','SiteLogo','SiteTemplateId','SiteTitle','Title','UniqueID','UniqueId','ViewCount','ViewsLifeTimeUniqueUsers','WebId','isDocument','isexternalcontent'],'Filter':null,'From':15,'Size':" + str(number_of_results) + ",'Query':{'DisplayQueryString': '" + searchterm + "','QueryString': '" + searchterm + "','QueryTemplate':'({searchterms})'},'Sort':[{'Field':'PersonalScore','SortDirection':'Desc'}],'EnableQueryUnderstanding':false,'EnableSpeller':false,'IdFormat':0,'ResultsMerge':{'Type':'Interleaved'},'RefiningQueries':null,'ExtendedQueries':[{'SearchProvider':'SharePoint','Query':{'Culture':1033,'EnableQueryRules':false,'TrimDuplicates':false,'BypassResultTypes':true,'ProcessBestBets':false,'ProcessPersonalFavorites':false,'EnableInterleaving':false,'EnableMultiGeo':true,'RankingModelId':'ABBAABBA-AAAA-AAAA-CCCC-000000000428'}}],'HitHighlight':{'HitHighlightedProperties':['HitHighlightedSummary'],'SummaryLength':200},'FederationContext':{'SpoFederationContext':{'UserContextUrl':'https://"+sub_domain+".sharepoint.com'}},'EnableResultAnnotations':true}],'Cvid':'86c342b5-0c27-1cff-da0c-3a9d4a403d5e','Culture':'en-us','Scenario':{'Dimensions':[{'DimensionName':'QueryType','DimensionValue':'AllResults'},{'DimensionName':'FormFactor','DimensionValue':'Web'}],'Name':'officehome'},'TextDecorations':'Off','TimeZone':'UTC','LogicalId':'200144d8-61e8-e98d-a7d4-a2a7a83b653e','QueryAlterationOptions':{'EnableSuggestion':true,'EnableAlteration':true,'SupportedRecourseDisplayTypes':['Suggestion','NoResultModification','NoRequeryModification']},'WholePageRankingOptions':{'EnableEnrichedRanking':true,'EnableLayoutHints':true,'SupportedSerpRegions':['MainLine'],'EntityResultTypeRankingOptions':[{'ResultType':'Answer','MaxEntitySetCount':6,'EntityTypeConstraints':[]}]}}"

    r = requests.post(url, data = query, headers = headers)
    if (r.status_code == 401):
        print ("Authorized / Access Denied code received.  Check your bearer token\n")
        exit(-1)
    if (r.status_code == 200):
        return (r.text)
    else:
        print("Unknown Response Received:\n")
        print(r.text)
        exit(-1)

def parse_results(raw_results):

    results_dict = json.loads(raw_results)
    results = results_dict["EntitySets"][0]["ResultSets"][0]["Results"]

    return results

def write_to_csv(results,searchterm):
 
    with open(searchterm + ".csv", 'w', newline = "\n") as f:
  
        writer = csv.DictWriter(f,results[0].keys())
        writer.writeheader()
        for r in results:
            writer = csv.DictWriter(f,r.keys())
            writer.writerow(r)  
def main():
    if len(sys.argv) > 3  or len(sys.argv) < 2:
        print("usage %s <search term> [number of results]\n" % sys.argv[0])
        exit(-1)
    if len(sys.argv) == 3:
        number_of_results = sys.argv[2]
    searchterm = sys.argv[1]

    raw_results = search(searchterm,number_of_results)
    results = parse_results(raw_results)
    write_to_csv(results,searchterm)
    exit()

if __name__ == "__main__":
    main()