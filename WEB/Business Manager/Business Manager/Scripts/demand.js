function DemandTrigger() {
    addMethod(this, "Trigger", function (Company, SecretKey, DemandNo, FlowNo, ItemNo) {
        let targetUrl = "/api/SSO/DemandFlowTrigger";
        let postData = {
            "Company": Company,
            "SecretKey": SecretKey,
            "DemandNo": DemandNo,
            "FlowNo": FlowNo,
            "ItemNo": ItemNo
        };

        let result = DataAccess.Update(targetUrl, postData, false, false);
    });

    addMethod(this, "Bind", function (Company, SecretKey, DemandNo, FlowNo, ItemNo) {
        let targetUrl = "/api/SSO/DemandItemBind";
        let postData = {
            "Company": Company,
            "SecretKey": SecretKey,
            "DemandNo": DemandNo,
            "FlowNo": FlowNo,
            "ItemNo": ItemNo
        };

        let result = DataAccess.Update(targetUrl, postData, false, false);
    });
}