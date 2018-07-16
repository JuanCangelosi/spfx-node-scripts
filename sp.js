const { Web, sp, FieldTypes, FieldCreationProperties } = require('@pnp/sp');
const { PnpNode } = require('sp-pnp-node');


new PnpNode().init().then(async (settings) => {

  sp.setup({
    sp: {
      headers: {
        'Accept': 'application/json;odata=verbose'
        // 'Accept': 'application/json;odata=minimalmetadata'
        // 'Accept': 'application/json;odata=nometadata'
      }
    }
  });

  // Here goes PnP JS Core code >>>

  const web = new Web(settings.siteUrl);
  let siteGroupManager;
  let siteGroupUser;
  try {
    siteGroupManager = await web.siteGroups.getByName("PACManager").get();
    siteGroupUser = await web.siteGroups.getByName("PACUser").get();
  } catch (error) {
    siteGroupManager = await web.siteGroups.add({
      Title: "PACManager",
      PrincipalType: 8,
      AllowMembersEditMembership: false,
      AllowRequestToJoinLeave: false,
      AutoAcceptRequestToJoinLeave: false,
      Description: "Members of this group can view pages, list items, and documents. If the document has a server rendering available, they can only view the document using the server rendering.",
      OnlyAllowMembersViewMembership: true
    })
    siteGroupUser = await web.siteGroups.add({
      Title: "PACUser",
      PrincipalType: 8,
      AllowMembersEditMembership: false,
      AllowRequestToJoinLeave: false,
      AutoAcceptRequestToJoinLeave: false,
      Description: "Members of this group can view pages, list items, and documents. If the document has a server rendering available, they can only view the document using the server rendering.",
      OnlyAllowMembersViewMembership: true
    })
  }


  // Get all content types example
  let listRequest;
  try {
    listRequest = web.lists.getByTitle("PACRequest");
    await listRequest.delete();
  } catch (error) {
    console.log("List Request Does not Exist");
  }
  const listRequestAddResult = await web.lists.add("PACRequest", "Permision Administration Control requests list", undefined, true);
  listRequest = listRequestAddResult.list;
  await listRequest.fields.addUser("PACRequestTo", 0, { SelectionGroup: siteGroupManager.Id });
  await listRequest.fields.addUser("PACNotify", 0, { SelectionGroup: siteGroupManager.Id, AllowMultipleValues: true });
  await listRequest.fields.addDateTime("PACDateFrom", 1);
  await listRequest.fields.addDateTime("PACDateTo", 1);
  await listRequest.fields.addBoolean("PACIsPeriod");
  await listRequest.fields.addChoice("PACRequestStatus", ['Pending', 'Approved', 'Denied']);
  await listRequest.fields.addMultilineText("PACReason", 5, false, false);
  await listRequest.fields.addChoice("PACRequestType", ['Vacation','Sickness', 'Travel', 'Paternity', 'Maternity', 'Medical Visit', 'Paperwork', 'Family Defunction', 'Other']);


  const listRequestProperties = await listRequest.get();
  const listRequestId = listRequestProperties.Id;

  let listResponse;
  try {
    listResponse = web.lists.getByTitle("PACResponse");
    await listResponse.delete();
  } catch (error) {
    console.log("List Response Does not Exist");
  }
  const listResponseAddResult = await web.lists.add("PACResponse", "Permision Administation Control responses list", undefined, true);
  listResponse = listResponseAddResult.list;
  await listResponse.fields.addLookup("PACRequest", listRequestId, "Title");
  await listResponse.fields.addChoice("PACResponse", ['Approved', 'Denied']);
  await listResponse.fields.addMultilineText("PACResponseReason", 5, false, false);


}).catch(console.log);