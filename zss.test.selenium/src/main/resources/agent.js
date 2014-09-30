{
if(window.zssts){return;}

var debug = zsstsDebug;

var tripId = (jq.now() % 9999) + 1;

var zssts= {};
window.zssts = zssts;

zssts._onResponse = function(){
	tripId++;
	if(tripId>9999){
		tripId = 0;
	}
};

zWatch.listen({
	onResponse : [zssts,zssts._onResponse]
});

zssts.getTripId = function(){
	return tripId;
}

zssts.jqSelectSingle = function(selector){
	var elm = jq(selector);
	if(debug){
		zk.log("jqSelectSingle "+selector+":"+elm);
	}
	
	if(elm.length>0){
		return elm[0];
	}
	return null;
}

if(debug){
	zk.log("zssts agent js loaded");
}
};