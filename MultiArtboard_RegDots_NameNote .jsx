#target illustrator
/*
 * MultiArtboard_RegDots_NameNote.jsx (v4 - order-aware, selective error focus)
 *
 * SOURCE: exactly one layer named Thru-cut (Thru cut / Thru_cut also accepted).
 * Each direct GroupItem, CompoundPathItem or PathItem is ONE panel. Nested
 * groups stay inside their outer group; sublayers do not implicitly group art.
 * Geometric bounds exclude stroke. Guides are ignored. Clipping groups and
 * non-path artwork in the cut layer are rejected instead of guessed at.
 *
 * FLOW (artwork is never moved):
 *   cut bounds -> width/height +1.00 inch -> Artboard A -> dots on A
 *              -> width/height +0.25 inch -> final artboard -> name / note.
 * Both expansions are TOTAL dimension increases, centered on the same point.
 * Thus A has 0.50 inch per-side clearance; final has 0.625 inch per side.
 * Existing artboards are replaced by these panels. A complete one-to-one
 * containment match automatically preserves the existing artboard order.
 * Otherwise a dialog offers existing-board order (spatial order within each
 * board) or spatial order for all panels. Empty old boards are skipped.
 * Spatial order is top-to-bottom rows, then left-to-right in each row.
 *
 * DOTS: 0.235 inch diameter, K100, no stroke, tangent INSIDE Artboard A.
 * Only ONE randomly selected corner gets a TOTAL budget of 1-2 extra dots,
 * shared between its two adjacent edges; insufficient space may reduce this.
 * Middle fill is irregular. Maximum actual along-edge center gap is 40 inch,
 * or 18 inch if A's width OR height exceeds 98 inch. Patterns are checked for
 * both-axis non-mirror-symmetry and uniqueness within this run.
 *
 * NOTE: final-board name and bottom label are P{index}- (random 8 digits).
 * Arial 18pt physical, M100; the label slides sideways to avoid the dots.
 *
 * PREFLIGHT: actual cut-path crossings/touches (including within groups),
 * self-crossings, and positive-area overlap/containment of A or final boards.
 * Bezier paths are adaptively flattened to 0.0005 inch. Separate paths within
 * about 0.002 inch are treated as a possible contact requiring inspection.
 * This is geometric preflight, not an exact symbolic Bezier intersection test.
 * Preflight failure changes no artboard and removes no old output.
 * Related top-level Groups / independent Shapes are selected on an overlap
 * or an invalid existing-board mapping. Their layers/ancestors are made
 * visible and unlocked as needed. No diagnostic artwork/layer is created.
 * New output is staged; runtime errors trigger restoration of old artboards.
 *
 * Successful runs replace the output layers Register and note. Original
 * Thru-cut / printing artwork is left as-is. Large-canvas constants use SF.
 * ExtendScript / ES3-compatible; no external dependencies.
 */

(function () {
    if (app.documents.length === 0) { alert("Open a document first."); return; }
    var doc = app.activeDocument;

    // ---- CONFIG: inches here are PHYSICAL inches ---------------------------
    var IN = 72, SF = doc.scaleFactor || 1;
    var FIRST_ADD_IN = 1.00;                // TOTAL extra width AND height
    var FINAL_ADD_IN = 0.25;                // additional TOTAL width AND height
    var DOT_DIA = 0.235 * IN / SF, DOT_R = DOT_DIA / 2;
    var GAP_STD = 40 * IN / SF, GAP_BIG = 18 * IN / SF;
    var BIG_EDGE = 98 * IN / SF;
    var FONT_PT = 18 / SF, FONT_NM = "ArialMT";
    var SLIDE_MIN = 0.4 * IN / SF, SLIDE_MAX = 1.0 * IN / SF;
    var EPS = 0.00001 * IN / SF;
    var FLAT_TOL = 0.0005 * IN / SF;
    var CONTACT_TOL = 0.002 * IN / SF;
    var ROW_ALIGN_TOL = 0.5 * IN / SF;      // top-edge tolerance for a spatial row
    var MAX_FLAT_SEGMENTS = 20000;          // per path; fail instead of truncating
    var MAX_ISSUES = 12;
    var CL_SEQS_IN = [
        [4], [4,4], [2,8], [8,2], [3,6], [6,3], [4,9], [9,4]
    ];

    var oldCS, changedCS = false;
    var regStage = null, noteStage = null, oldReg = null, oldNote = null;
    var oldBoards = null, oldActive = 0, boardsTouched = false, committed = false;
    var regState = null, noteState = null, fontWarned = false;
    var warnings = [];
    var problemUnits = [];

    function fail(msg) { throw new Error(msg); }
    function trim(s) { return String(s).replace(/^\s+|\s+$/g, ""); }
    function normalizeLayer(s) { return trim(s).toLowerCase().replace(/[\s_\-\u2010-\u2015]+/g, ""); }
    function clamp(v, lo, hi) { return v < lo ? lo : (v > hi ? hi : v); }
    function sq(v) { return v*v; }
    function distance(a, b) { return Math.sqrt(sq(a.x-b.x)+sq(a.y-b.y)); }
    function point(a) { return {x:a[0], y:a[1]}; }
    function midpoint(a, b) { return {x:(a.x+b.x)/2, y:(a.y+b.y)/2}; }
    function expandRect(r, total) {
        var p = total/2;
        return [r[0]-p, r[1]+p, r[2]+p, r[3]-p];
    }
    function boundsOverlap(a, b, tol) {
        return Math.min(a[2],b[2]) >= Math.max(a[0],b[0])-tol &&
               Math.min(a[1],b[1]) >= Math.max(a[3],b[3])-tol;
    }
    function boardOverlap(a, b) {
        // Merely sharing an edge is allowed; overlapping/contained areas are not.
        return Math.min(a[2],b[2])-Math.max(a[0],b[0]) > EPS &&
               Math.min(a[1],b[1])-Math.max(a[3],b[3]) > EPS;
    }
    function validBounds(r, label) {
        if (!r || r.length !== 4) fail("Cannot read bounds: " + label);
        for (var i=0; i<4; i++) if (!isFinite(r[i])) fail("Invalid bounds: " + label);
        if (r[2] < r[0] || r[1] < r[3]) fail("Invalid bounds: " + label);
        if (r[2]-r[0] <= EPS && r[1]-r[3] <= EPS) fail("Zero-size cut object: " + label);
    }
    function mkCMYK(c,m,y,k) {
        var col = new CMYKColor();
        col.cyan=c; col.magenta=m; col.yellow=y; col.black=k; return col;
    }
    function describe(item, index) {
        var name = trim(item.name || "");
        return item.typename + " " + (index+1) + (name ? ' "'+name+'"' : "");
    }

    // ---- Thru-cut collection: filter direct parents to avoid double-counting
    function findCutLayers(layers, out) {
        for (var i=0; i<layers.length; i++) {
            var L = layers[i];
            if (normalizeLayer(L.name) === "thrucut") out.push(L);
            else findCutLayers(L.layers, out);
        }
    }
    function collectPaths(item, out, label) {
        if (item.typename === "PathItem" && item.guides) return;
        if (item.hidden) fail("Hidden cut object: "+label+". Show it or remove it from Thru-cut.");
        var i, child;
        if (item.typename === "PathItem") {
            if (item.clipping) fail("Clipping path in Thru-cut: "+label);
            if (item.pathPoints.length < 2) fail("Cut path has fewer than 2 anchors: "+label);
            out.push({item:item, label:label});
        } else if (item.typename === "CompoundPathItem") {
            for (i=0; i<item.pathItems.length; i++) {
                collectPaths(item.pathItems[i], out, label+" / subpath "+(i+1));
            }
        } else if (item.typename === "GroupItem") {
            if (item.clipped) fail("Clipping group in Thru-cut: "+label+". Use ordinary cut-path groups.");
            for (i=0; i<item.pageItems.length; i++) {
                child=item.pageItems[i];
                if (child.parent === item) collectPaths(child, out, label+" / "+describe(child,i));
            }
        } else {
            fail("Unsupported object in Thru-cut: "+label+". Keep only paths and ordinary groups there.");
        }
    }
    function collectUnits(L, units) {
        for (var i=0; i<L.pageItems.length; i++) {
            var item=L.pageItems[i];
            if (item.parent !== L) continue;
            var paths=[], label=describe(item,units.length);
            collectPaths(item, paths, label);
            if (!paths.length) continue; // guide-only object/group
            // Union real path bounds: identical to a path-only group's bounds,
            // while excluding guide geometry and stroke thickness.
            var r=null;
            for (var j=0; j<paths.length; j++) {
                var b=paths[j].item.geometricBounds;
                validBounds(b, paths[j].label);
                if (!r) r=[b[0],b[1],b[2],b[3]];
                else r=[Math.min(r[0],b[0]),Math.max(r[1],b[1]),Math.max(r[2],b[2]),Math.min(r[3],b[3])];
            }
            var a=expandRect(r,FIRST_ADD_IN*IN/SF);
            units.push({item:item, label:label, paths:paths, cutBounds:r,
                        sourceIndex:units.length, rectA:a,
                        rectFinal:expandRect(a,FINAL_ADD_IN*IN/SF)});
        }
        for (var k=0; k<L.layers.length; k++) collectUnits(L.layers[k], units);
    }

    // ---- Numbering: complete automatic match, otherwise an explicit choice -
    function containsRect(outer,inner) {
        return inner[0]>=outer[0]-EPS && inner[2]<=outer[2]+EPS &&
               inner[1]<=outer[1]+EPS && inner[3]>=outer[3]-EPS;
    }
    function matchExistingBoards(units) {
        var rects=[],counts=[],invalid=[],matched=0,used=0;
        for(var i=0;i<doc.artboards.length;i++) {rects.push(doc.artboards[i].artboardRect);counts.push(0);}
        for(var u=0;u<units.length;u++) {
            var candidates=[];
            for(var b=0;b<rects.length;b++) if(containsRect(rects[b],units[u].cutBounds)) candidates.push(b);
            units[u].oldBoardIndex=candidates.length===1 ? candidates[0] : -1;
            if(candidates.length===1) {counts[candidates[0]]++;matched++;}
            else invalid.push({unit:units[u],reason:candidates.length ?
                "Fits multiple existing artboards" : "Crosses an existing board edge or is outside the boards"});
        }
        var exact=matched===units.length && rects.length===units.length;
        for(var c=0;c<counts.length;c++) {
            if(counts[c]>0)used++;
            if(counts[c]!==1)exact=false;
        }
        return {exact:exact,counts:counts,invalid:invalid,used:used};
    }
    function spatialSort(list) {
        var pending=list.slice(),out=[],row=[],anchorY=null;
        pending.sort(function(a,b){
            return b.cutBounds[1]-a.cutBounds[1] || a.cutBounds[0]-b.cutBounds[0] || a.sourceIndex-b.sourceIndex;
        });
        function flush() {
            row.sort(function(a,b){
                return a.cutBounds[0]-b.cutBounds[0] || b.cutBounds[1]-a.cutBounds[1] || a.sourceIndex-b.sourceIndex;
            });
            for(var j=0;j<row.length;j++)out.push(row[j]);
            row=[];
        }
        // A fixed row anchor avoids a chain of slightly shifted tops merging
        // several rows together, and avoids a non-transitive sort comparator.
        for(var i=0;i<pending.length;i++) {
            var y=pending[i].cutBounds[1];
            if(anchorY===null || anchorY-y>ROW_ALIGN_TOL+EPS) {
                flush();anchorY=y;
            }
            row.push(pending[i]);
        }
        flush();return out;
    }
    function chooseOrderMode(units,matches) {
        if(typeof Window==="undefined") fail("Panel numbering needs an Illustrator ScriptUI dialog. No changes were made.");
        var win=new Window("dialog","Panel numbering");
        win.orientation="column";win.alignChildren="fill";
        win.add("statictext",undefined,"How should the new artboards be numbered?");
        win.add("statictext",undefined,units.length+" cut panels / "+doc.artboards.length+" current artboards. Layout is not one-to-one.");
        var choices=win.add("group");choices.orientation="column";choices.alignChildren="left";
        var keep=choices.add("radiobutton",undefined,"Use existing artboard order");
        var position=choices.add("radiobutton",undefined,"Sort all panels by position");
        keep.value=matches.invalid.length===0 && matches.used>1;position.value=!keep.value;
        var help=win.add("statictext",undefined,"",{multiline:true});help.preferredSize=[470,80];
        function explain() {
            help.text=keep.value ?
                "Existing board order first. Multiple panels inside one board are ordered top-to-bottom, then left-to-right.\nEmpty boards are skipped; new P numbers are consecutive.\nPanels without one clear original board will be selected and the script will stop." :
                "Existing board order is ignored. All cut panels are ordered top-to-bottom, then left-to-right within each row.";
        }
        keep.onClick=explain;position.onClick=explain;explain();
        var buttons=win.add("group");buttons.alignment="right";
        var go=buttons.add("button",undefined,"Continue",{name:"ok"});
        var cancel=buttons.add("button",undefined,"Cancel",{name:"cancel"});
        win.defaultElement=go;win.cancelElement=cancel;
        if(win.show()!==1)return null;
        return keep.value ? "keep" : "position";
    }
    function orderUnits(units) {
        var matches=matchExistingBoards(units);
        var mode=matches.exact ? "keep" : chooseOrderMode(units,matches);
        if(mode===null)return null;
        if(mode==="position")return {units:spatialSort(units),description:"Position order (top-to-bottom rows, left-to-right)"};
        if(matches.invalid.length) {
            var messages=[];
            for(var i=0;i<matches.invalid.length;i++) {
                var entry=matches.invalid[i];markProblem(entry.unit);
                if(messages.length<MAX_ISSUES)messages.push(entry.reason+": "+entry.unit.label);
            }
            fail("Cannot preserve the existing artboard order. Check the selected cut objects, or rerun and choose position order.\n\n"+messages.join("\n\n"));
        }
        var sorted=[];
        for(var b=0;b<matches.counts.length;b++) {
            var bucket=[];
            for(var u=0;u<units.length;u++)if(units[u].oldBoardIndex===b)bucket.push(units[u]);
            bucket=spatialSort(bucket);
            for(var p=0;p<bucket.length;p++)sorted.push(bucket[p]);
        }
        return {units:sorted,description:matches.exact ?
            "Existing artboard order (automatic one-to-one match)" :
            "Existing board order; position order within each board"};
    }

    // ---- Error focus: select real cut Groups / Shapes, never add artwork ----
    function markProblem(unit) {
        if(!unit)return;
        for(var i=0;i<problemUnits.length;i++)if(problemUnits[i].item===unit.item)return;
        problemUnits.push(unit);
    }
    function makeSelectable(item,state) {
        var chain=[],p=item;
        while(p && p.typename!=="Document") {chain.push(p);p=p.parent;}
        // Unlock/show outer ancestors before touching their descendants.
        for(var i=chain.length-1;i>=0;i--) {
            p=chain[i];
            if(p.locked) {p.locked=false;state.changed=true;}
            if(p.typename==="Layer") {
                if(!p.visible) {p.visible=true;state.changed=true;}
            } else if(p.hidden) {p.hidden=false;state.changed=true;}
        }
    }
    function selectProblems() {
        var state={changed:false},notes=[],selected=0;
        try{doc.selection=null;}catch(clearErr){notes.push("Could not clear the previous selection.");}
        for(var i=0;i<problemUnits.length;i++) {
            var unit=problemUnits[i];
            try {
                for(var p=0;p<unit.paths.length;p++)makeSelectable(unit.paths[p].item,state);
                makeSelectable(unit.item,state);
                unit.item.selected=true;
                if(!unit.item.selected)throw new Error("Object remained unselected");
                selected++;
            } catch(selectErr) {notes.push("Could not select "+unit.label+": "+selectErr.message);}
        }
        notes.unshift("Related Groups / Shapes selected: "+selected);
        if(state.changed)notes.push("Related objects/layers were made visible and unlocked for inspection.");
        try{app.redraw();}catch(redrawErr){}
        return notes.join("\n");
    }

    // ---- Path preflight: adaptive cubic subdivision + segment intersection --
    function pointSegmentDistance(p, a, b) {
        var dx=b.x-a.x, dy=b.y-a.y, len=dx*dx+dy*dy;
        if (len <= EPS*EPS) return distance(p,a);
        var t=clamp(((p.x-a.x)*dx+(p.y-a.y)*dy)/len,0,1);
        return distance(p,{x:a.x+t*dx,y:a.y+t*dy});
    }
    function flattenCubic(a, b, c, d, out, depth, label) {
        // Distance to the finite chord also detects collinear backtracking.
        if (Math.max(pointSegmentDistance(b,a,d),pointSegmentDistance(c,a,d)) <= FLAT_TOL) {
            out.push(d);
            if (out.length > MAX_FLAT_SEGMENTS+1) fail("Cut path too complex to check: "+label);
            return;
        }
        if (depth >= 24) fail("Could not resolve cut curve accurately: "+label);
        var ab=midpoint(a,b), bc=midpoint(b,c), cd=midpoint(c,d);
        var abc=midpoint(ab,bc), bcd=midpoint(bc,cd), mid=midpoint(abc,bcd);
        flattenCubic(a,ab,abc,mid,out,depth+1,label);
        flattenCubic(mid,bcd,cd,d,out,depth+1,label);
    }
    function flattenPath(entry) {
        var P=entry.item, pts=P.pathPoints, vertices=[point(pts[0].anchor)];
        var count=P.closed ? pts.length : pts.length-1;
        for (var i=0; i<count; i++) {
            var next=(i+1)%pts.length;
            flattenCubic(point(pts[i].anchor),point(pts[i].rightDirection),
                         point(pts[next].leftDirection),point(pts[next].anchor),vertices,0,entry.label);
        }
        var segments=[];
        for (var j=1; j<vertices.length; j++) {
            var a=vertices[j-1], b=vertices[j];
            if (distance(a,b) <= EPS) continue;
            segments.push({a:a,b:b,bounds:[Math.min(a.x,b.x),Math.max(a.y,b.y),
                                         Math.max(a.x,b.x),Math.min(a.y,b.y)],index:segments.length});
        }
        if (!segments.length) fail("Zero-length cut path: "+entry.label);
        var sorted=segments.slice();
        sorted.sort(function(a,b){return a.bounds[0]-b.bounds[0];});
        return {segments:segments, sorted:sorted, closed:P.closed,
                bounds:P.geometricBounds, label:entry.label};
    }
    function cross(a,b,c) { return (b.x-a.x)*(c.y-a.y)-(b.y-a.y)*(c.x-a.x); }
    function segmentsCross(a,b,c,d) {
        var abC=cross(a,b,c), abD=cross(a,b,d), cdA=cross(c,d,a), cdB=cross(c,d,b);
        return ((abC>0 && abD<0)||(abC<0 && abD>0)) &&
               ((cdA>0 && cdB<0)||(cdA<0 && cdB>0));
    }
    function segmentsContact(s,t,tol) {
        if (!boundsOverlap(s.bounds,t.bounds,tol)) return false;
        return segmentsCross(s.a,s.b,t.a,t.b) ||
               pointSegmentDistance(s.a,t.a,t.b) <= tol ||
               pointSegmentDistance(s.b,t.a,t.b) <= tol ||
               pointSegmentDistance(t.a,s.a,s.b) <= tol ||
               pointSegmentDistance(t.b,s.a,s.b) <= tol;
    }
    function selfCrosses(P) {
        var s=P.sorted, n=P.segments.length;
        for (var i=0; i<s.length; i++) for (var j=i+1; j<s.length; j++) {
            if (s[j].bounds[0] > s[i].bounds[2]+EPS) break;
            var diff=Math.abs(s[i].index-s[j].index);
            if (diff===1 || (P.closed && diff===n-1)) {
                // Adjacent segments meet normally, but reversing along the
                // same segment is still an overlapping cut.
                var a=s[i], b=s[j], shared=null, oa, ob;
                if (distance(a.b,b.a)<=EPS) { shared=a.b; oa=a.a; ob=b.b; }
                else if (distance(b.b,a.a)<=EPS) { shared=a.a; oa=a.b; ob=b.a; }
                if (shared && distance(oa,shared)>EPS && distance(ob,shared)>EPS &&
                    Math.abs(cross(shared,oa,ob)) <= EPS*Math.max(distance(oa,shared),distance(ob,shared)) &&
                    (oa.x-shared.x)*(ob.x-shared.x)+(oa.y-shared.y)*(ob.y-shared.y)>0) return true;
                continue;
            }
            // EPS for self-test: neighboring tiny curve chords must not be
            // mistaken for contact merely because they're within CONTACT_TOL.
            if (segmentsContact(s[i],s[j],EPS)) return true;
        }
        return false;
    }
    function pathsContact(A,B) {
        if (!boundsOverlap(A.bounds,B.bounds,CONTACT_TOL)) return false;
        var a=A.sorted, b=B.sorted, start=0;
        for (var i=0; i<a.length; i++) {
            while (start<b.length && b[start].bounds[2]<a[i].bounds[0]-CONTACT_TOL) start++;
            for (var j=start; j<b.length; j++) {
                if (b[j].bounds[0]>a[i].bounds[2]+CONTACT_TOL) break;
                if (segmentsContact(a[i],b[j],CONTACT_TOL)) return true;
            }
        }
        return false;
    }
    function preflight(units) {
        var issues=[], all=[];
        function report(s,unitA,unitB) {
            // The displayed text is capped, but every implicated panel is kept
            // for selection, including issues beyond the first twelve.
            markProblem(unitA);markProblem(unitB);
            if (issues.length<MAX_ISSUES) issues.push(s);
        }
        for (var i=0; i<units.length; i++) {
            for (var j=0; j<units[i].paths.length; j++) {
                var p=flattenPath(units[i].paths[j]);p.unit=units[i];all.push(p);
                if (selfCrosses(p)) report("Self-crossing/overlapping path: "+p.label,p.unit);
            }
        }
        for (var a=0; a<all.length; a++) for (var b=a+1; b<all.length; b++) {
            if (pathsContact(all[a],all[b])) report("Crossing/touching cut paths: "+all[a].label+" <> "+all[b].label,all[a].unit,all[b].unit);
        }
        for (var u=0; u<units.length; u++) for (var v=u+1; v<units.length; v++) {
            if (boardOverlap(units[u].rectA,units[v].rectA)) {
                report("Artboard A overlap: "+units[u].label+" <> "+units[v].label,units[u],units[v]);
            } else if (boardOverlap(units[u].rectFinal,units[v].rectFinal)) {
                report("Final artboard overlap (+0.25 inch): "+units[u].label+" <> "+units[v].label,units[u],units[v]);
            }
        }
        if (issues.length) fail("Check the Thru-cut artwork before running again.\n\n"+
                               issues.join("\n\n")+(issues.length>=MAX_ISSUES ? "\n\n(First 12 issues shown.)" : ""));
    }

    // ---- Dots: one corner shares a TOTAL 1-2-dot budget between both arms ---
    function pickSeq(count) {
        if(count===0)return [];
        var menu=[];
        for(var m=0;m<CL_SEQS_IN.length;m++)if(CL_SEQS_IN[m].length===count)menu.push(CL_SEQS_IN[m]);
        var src=menu[Math.floor(Math.random()*menu.length)],out=[];
        // Small continuous jitter keeps layouts distinguishable even for many
        // equal small panels, where no middle fill is needed.
        for(var i=0;i<src.length;i++)out.push((src[i]+(Math.random()-0.5)*0.5)*IN/SF);
        return out;
    }
    function cornerPlan() {
        var total=1+Math.floor(Math.random()*2),hCount=0,vCount=0;
        if(total===1) {
            if(Math.random()<0.5)hCount=1;else vCount=1;
        } else {
            hCount=Math.floor(Math.random()*3);vCount=2-hCount;
        }
        return {horizontal:pickSeq(hCount),vertical:pickSeq(vCount),budget:total};
    }
    function edgeDots(lo,hi,maxGap,loSeq,hiSeq) {
        var U=hi-lo;
        if (U<=EPS) return [(lo+hi)/2];
        var out=[lo], cap=Math.min(U*0.45,16*IN/SF), minSep=1*IN/SF;
        var p=lo;
        for (var i=0;i<loSeq.length;i++) {
            p+=loSeq[i];
            if (p-lo<=cap && p<hi-minSep) out.push(p); else break;
        }
        var loInner=out[out.length-1];
        var hiPts=[hi], q=hi;
        for (var j=0;j<hiSeq.length;j++) {
            q-=hiSeq[j];
            if (hi-q<=cap && q>loInner+minSep) hiPts.push(q); else break;
        }
        var hiInner=hiPts[hiPts.length-1], span=hiInner-loInner;
        if (span>maxGap) {
            var n=Math.ceil(span/(maxGap*0.7)), base=span/n;
            // +/-0.18 base -> adjacent worst case <= 1.36*0.7 = 0.952 maxGap.
            for (var m=1;m<n;m++) out.push(loInner+m*base+(Math.random()-0.5)*base*0.36);
        }
        for (var h=0;h<hiPts.length;h++) out.push(hiPts[h]);
        return out;
    }
    function dedupe(dots) {
        var out=[];
        for (var i=0;i<dots.length;i++) {
            var keep=true;
            for (var j=0;j<out.length;j++) if (distance(dots[i],out[j])<DOT_DIA-EPS) {keep=false;break;}
            if (keep) out.push(dots[i]);
        }
        return out;
    }
    function rk(v) {return Math.round(v/EPS);}
    function mirrorSym(dots,cx,cy) {
        // Relative coordinates avoid asymmetric rounding at global offsets.
        var keys={}, v=true,h=true;
        for (var i=0;i<dots.length;i++) keys[rk(dots[i].x-cx)+","+rk(dots[i].y-cy)]=true;
        for (var j=0;j<dots.length;j++) {
            var x=rk(dots[j].x-cx),y=rk(dots[j].y-cy);
            if (!keys[(-x)+","+y]) v=false;
            if (!keys[x+","+(-y)]) h=false;
        }
        return {v:v,h:h};
    }
    function slideAlongEdge(d,r) {
        var amt=(Math.random()<0.5 ? -1 : 1)*(SLIDE_MIN+Math.random()*(SLIDE_MAX-SLIDE_MIN));
        if (Math.abs(d.y-(r[1]-DOT_R))<EPS || Math.abs(d.y-(r[3]+DOT_R))<EPS)
            d.x=clamp(d.x+amt,r[0]+DOT_R,r[2]-DOT_R);
        else d.y=clamp(d.y+amt,r[3]+DOT_R,r[1]-DOT_R);
    }
    function localKey(dots,r) {
        var arr=[];
        for (var i=0;i<dots.length;i++) arr.push(rk(dots[i].x-r[0])+":"+rk(dots[i].y-r[3]));
        arr.sort();return arr.join("|");
    }
    function edgePoints(dots,r,edge) {
        var arr=[];
        for (var i=0;i<dots.length;i++) {
            var d=dots[i];
            if (edge===0 && Math.abs(d.y-(r[1]-DOT_R))<EPS) arr.push(d.x);
            if (edge===1 && Math.abs(d.y-(r[3]+DOT_R))<EPS) arr.push(d.x);
            if (edge===2 && Math.abs(d.x-(r[0]+DOT_R))<EPS) arr.push(d.y);
            if (edge===3 && Math.abs(d.x-(r[2]-DOT_R))<EPS) arr.push(d.y);
        }
        arr.sort(function(a,b){return a-b;});return arr;
    }
    function repairGaps(dots,r,maxGap) {
        for (var edge=0;edge<4;edge++) {
            var arr=edgePoints(dots,r,edge);
            for (var i=1;i<arr.length;i++) {
                var gap=arr[i]-arr[i-1];
                if (gap<=maxGap+EPS) continue;
                var n=Math.ceil(gap/(maxGap*0.9));
                for (var j=1;j<n;j++) {
                    var value=arr[i-1]+gap*j/n;
                    dots.push(edge<2 ? {x:value,y:edge===0?r[1]-DOT_R:r[3]+DOT_R} :
                                      {x:edge===2?r[0]+DOT_R:r[2]-DOT_R,y:value});
                }
            }
        }
        return dedupe(dots);
    }
    function validateDots(dots,r,maxGap) {
        for (var i=0;i<dots.length;i++) {
            var d=dots[i];
            if (d.x<r[0]+DOT_R-EPS || d.x>r[2]-DOT_R+EPS ||
                d.y<r[3]+DOT_R-EPS || d.y>r[1]-DOT_R+EPS) return false;
            for (var j=0;j<i;j++) if (distance(d,dots[j])<DOT_DIA-EPS) return false;
        }
        for (var edge=0;edge<4;edge++) {
            var arr=edgePoints(dots,r,edge);
            for (var k=1;k<arr.length;k++) if (arr[k]-arr[k-1]>maxGap+EPS) return false;
        }
        return dots.length>=3;
    }
    function buildDots(r,seen,maxGap) {
        var cx=(r[0]+r[2])/2,cy=(r[1]+r[3])/2;
        for (var attempt=0;attempt<64;attempt++) {
            // 0=top-left, 1=top-right, 2=bottom-left, 3=bottom-right.
            var corner=Math.floor(Math.random()*4),plan=cornerPlan(),raw=[],a,k;
            a=edgeDots(r[0]+DOT_R,r[2]-DOT_R,maxGap,corner===0?plan.horizontal:[],corner===1?plan.horizontal:[]);
            for(k=0;k<a.length;k++) raw.push({x:a[k],y:r[1]-DOT_R});
            a=edgeDots(r[0]+DOT_R,r[2]-DOT_R,maxGap,corner===2?plan.horizontal:[],corner===3?plan.horizontal:[]);
            for(k=0;k<a.length;k++) raw.push({x:a[k],y:r[3]+DOT_R});
            a=edgeDots(r[3]+DOT_R,r[1]-DOT_R,maxGap,corner===2?plan.vertical:[],corner===0?plan.vertical:[]);
            for(k=0;k<a.length;k++) raw.push({x:r[0]+DOT_R,y:a[k]});
            a=edgeDots(r[3]+DOT_R,r[1]-DOT_R,maxGap,corner===3?plan.vertical:[],corner===1?plan.vertical:[]);
            for(k=0;k<a.length;k++) raw.push({x:r[2]-DOT_R,y:a[k]});
            var dots=dedupe(raw);
            if (dots.length<=4) {
                dots.splice(Math.floor(Math.random()*dots.length),1);
                slideAlongEdge(dots[Math.floor(Math.random()*dots.length)],r);
            }
            for (var guard=0;guard<12;guard++) {
                var sym=mirrorSym(dots,cx,cy);
                if (!sym.v && !sym.h) break;
                slideAlongEdge(dots[Math.floor(Math.random()*dots.length)],r);
            }
            dots=repairGaps(dedupe(dots),r,maxGap);
            var finalSym=mirrorSym(dots,cx,cy),key=localKey(dots,r);
            if (!finalSym.v && !finalSym.h && !seen[key] && validateDots(dots,r,maxGap)) {
                seen[key]=true;return dots;
            }
        }
        fail("Could not generate a distinct, non-symmetric dot pattern for a panel.");
    }

    // ---- Label collision: nearest horizontal clear interval ----------------
    function findClearX(cx,reach,xs,lo,hi) {
        if (lo>hi) return null;
        var ints=[];
        for (var i=0;i<xs.length;i++) {
            var a=xs[i]-reach,b=xs[i]+reach;
            if (b<lo || a>hi) continue;
            ints.push([Math.max(a,lo),Math.min(b,hi)]);
        }
        if (!ints.length) return clamp(cx,lo,hi);
        ints.sort(function(p,q){return p[0]-q[0];});
        var merged=[ints[0].slice()];
        for (var j=1;j<ints.length;j++) {
            var last=merged[merged.length-1];
            if (ints[j][0]<=last[1]+EPS) last[1]=Math.max(last[1],ints[j][1]);
            else merged.push(ints[j].slice());
        }
        var best=null,bestD=1e100,cur=lo;
        function consider(a,b) {
            if (b-a<=EPS) return;
            var x=clamp(cx,a,b),d=Math.abs(x-cx);
            if(d<bestD){best=x;bestD=d;}
        }
        for (var k=0;k<merged.length;k++) {
            if (merged[k][0]>cur+EPS) consider(cur,merged[k][0]-EPS);
            cur=Math.max(cur,merged[k][1]+EPS);
        }
        if (cur<hi-EPS) consider(cur,hi);
        return best;
    }
    function makeNote(unit,theFont,textColor) {
        var r=unit.rectFinal,cx=(r[0]+r[2])/2;
        var tf=noteStage.textFrames.pointText([cx,r[3]]);tf.contents=unit.name;
        var ca=tf.textRange.characterAttributes;
        ca.size=FONT_PT;ca.fillColor=textColor;
        if(theFont) ca.textFont=theFont;else fontWarned=true;
        tf.textRange.paragraphAttributes.justification=Justification.CENTER;
        var gb=tf.geometricBounds;
        tf.translate(cx-(gb[0]+gb[2])/2,r[3]+EPS-gb[3]);gb=tf.geometricBounds;
        var halfW=(gb[2]-gb[0])/2;
        if (gb[1]>r[1]-EPS) fail("Note is taller than its artboard: "+unit.label);
        var reach=halfW+DOT_R+DOT_DIA*0.4,xs=[];
        for (var i=0;i<unit.dots.length;i++) {
            var d=unit.dots[i];
            if(d.y+DOT_R>=gb[3] && d.y-DOT_R<=gb[1]) xs.push(d.x);
        }
        var x=findClearX(cx,reach,xs,r[0]+halfW+EPS,r[2]-halfW-EPS);
        if(x===null) fail("No room for an 18pt note without touching dots: "+unit.label+
                         ". Increase the panel margin or reduce FONT_PT in CONFIG.");
        tf.translate(x-cx,0);
    }

    // ---- Staging / reversible artboard replacement -------------------------
    function outputLayer(name) {
        var found=null;
        for(var i=0;i<doc.layers.length;i++) if(doc.layers[i].name===name) {
            if(found) fail("More than one top-level layer named "+name+". Rename duplicates first.");
            found=doc.layers[i];
        }
        return found;
    }
    function layerState(L) {return L ? {name:L.name,visible:L.visible,locked:L.locked} : null;}
    function isAncestor(L,item) {
        if(!L)return false;
        var p=item;
        while(p && p.typename!=="Document") {if(p===L)return true;p=p.parent;}
        return false;
    }
    function snapshotBoards() {
        var out=[];
        for(var i=0;i<doc.artboards.length;i++) {
            var a=doc.artboards[i];
            out.push({rect:a.artboardRect.slice(),name:a.name,ruler:a.rulerOrigin.slice(),
                      center:a.showCenterMark,cross:a.showCrossHairs,safe:a.showSafeAreas});
        }
        return out;
    }
    function restoreBoards(saved) {
        while(doc.artboards.length<saved.length) doc.artboards.add(saved[doc.artboards.length].rect);
        for(var i=0;i<saved.length;i++) {
            var a=doc.artboards[i],s=saved[i];a.artboardRect=s.rect;a.name=s.name;
            a.rulerOrigin=s.ruler;a.showCenterMark=s.center;a.showCrossHairs=s.cross;a.showSafeAreas=s.safe;
        }
        while(doc.artboards.length>saved.length) doc.artboards.remove(doc.artboards.length-1);
        doc.artboards.setActiveArtboardIndex(Math.min(oldActive,saved.length-1));
    }
    function restoreLayer(L,state) {
        if(L && state) {L.locked=false;L.name=state.name;L.visible=state.visible;L.locked=state.locked;}
    }
    function removeStage(L) {if(L){L.locked=false;L.remove();}}
    function stageLayer(name) {var L=doc.layers.add();L.name=name;return L;}
    function gen8() {var s="";for(var i=0;i<8;i++)s+=String(Math.floor(Math.random()*10));return s;}
    function uniqueId8(used) {var s;do{s=gen8();}while(used[s]);used[s]=true;return s;}

    try {
        if(typeof CoordinateSystem!=="undefined") {
            oldCS=app.coordinateSystem;app.coordinateSystem=CoordinateSystem.DOCUMENTCOORDINATESYSTEM;changedCS=true;
        }
        var cutLayers=[];findCutLayers(doc.layers,cutLayers);
        if(cutLayers.length!==1) fail(cutLayers.length ?
            "Found multiple Thru-cut layers. Keep one unambiguous cut layer." :
            "No Thru-cut layer found. Put cut paths in a layer named Thru-cut.");
        var units=[];collectUnits(cutLayers[0],units);
        if(!units.length) fail("Thru-cut contains no usable cut paths.");
        oldReg=outputLayer("Register");oldNote=outputLayer("note");
        if(isAncestor(oldReg,cutLayers[0]) || isAncestor(oldNote,cutLayers[0]))
            fail("Thru-cut must not be inside Register or note: those are generated output layers.");
        preflight(units);
        var numbering=orderUnits(units);
        if(numbering===null)return;          // user cancelled; no document edits
        units=numbering.units;

        var seen={},used={},totalDots=0;
        for(var u=0;u<units.length;u++) {
            var r=units[u].rectA,maxGap=(r[2]-r[0]>BIG_EDGE || r[1]-r[3]>BIG_EDGE)?GAP_BIG:GAP_STD;
            units[u].dots=buildDots(r,seen,maxGap);
            units[u].name="P"+(u+1)+"- ("+uniqueId8(used)+")";
            totalDots+=units[u].dots.length;
        }
        var theFont=null;
        try{theFont=app.textFonts.getByName(FONT_NM);}catch(f1){try{theFont=app.textFonts.getByName("Arial");}catch(f2){}}
        var dotColor=mkCMYK(0,0,0,100),textColor=mkCMYK(0,100,0,0);
        var token=String(new Date().getTime());
        regStage=stageLayer("__Register_build_"+token);
        noteStage=stageLayer("__note_build_"+token);
        for(var n=0;n<units.length;n++) {
            var dots=units[n].dots;
            for(var d=0;d<dots.length;d++) {
                var p=dots[d],el=regStage.pathItems.ellipse(p.y+DOT_R,p.x-DOT_R,DOT_DIA,DOT_DIA);
                el.stroked=false;el.filled=true;el.fillColor=dotColor;el.fillOverprint=false;
            }
            makeNote(units[n],theFont,textColor);
        }

        oldActive=doc.artboards.getActiveArtboardIndex();oldBoards=snapshotBoards();
        boardsTouched=true;
        // Reuse existing indices so no temporary artboard-count doubling occurs.
        for(var a=0;a<units.length;a++) {
            if(a<doc.artboards.length) doc.artboards[a].artboardRect=units[a].rectA;
            else doc.artboards.add(units[a].rectA);
        }
        while(doc.artboards.length>units.length) doc.artboards.remove(doc.artboards.length-1);
        // Dots stay on A; only the actual artboard expands for the final margin.
        for(var b=0;b<units.length;b++) {
            var ab=doc.artboards[b];ab.artboardRect=units[b].rectFinal;ab.name=units[b].name;
            ab.rulerOrigin=[0,0];ab.showCenterMark=false;ab.showCrossHairs=false;ab.showSafeAreas=false;
        }
        doc.artboards.setActiveArtboardIndex(0);

        regState=layerState(oldReg);noteState=layerState(oldNote);
        if(oldReg){oldReg.locked=false;oldReg.name="__Register_previous_"+token;oldReg.visible=false;}
        if(oldNote){oldNote.locked=false;oldNote.name="__note_previous_"+token;oldNote.visible=false;}
        regStage.name="Register";noteStage.name="note";
        committed=true;
        // Cleanup after commit: an old locked descendant must not invalidate
        // otherwise finished output; a retained backup is named explicitly.
        if(oldReg)try{oldReg.remove();}catch(c1){warnings.push("Old Register backup retained: "+oldReg.name);}
        if(oldNote)try{oldNote.remove();}catch(c2){warnings.push("Old note backup retained: "+oldNote.name);}
        if(fontWarned)warnings.push("Arial unavailable: note uses Illustrator's default font.");
        if(warnings.length)warnings.unshift("Warnings:");
        alert("Done.\nPanels / artboards: "+units.length+"\nRegister dots: "+totalDots+
              "\nNumbering: "+numbering.description+
              "\nCorner extras: 1-2 total, shared between both adjacent edges"+
              "\nA: cut width/height +1.00 inch\nFinal: cut width/height +1.25 inch"+
              "\nMax dot gap: 40 inch / 18 inch (A >98 inch)\nscaleFactor: "+SF+
              (warnings.length ? "\n\n"+warnings.join("\n") : ""));
    } catch(err) {
        var recovery=[],focus="";
        if(!committed) {
            if(boardsTouched && oldBoards)try{restoreBoards(oldBoards);}catch(r1){recovery.push("Artboard restore failed: "+r1.message);}
            try{restoreLayer(oldReg,regState);}catch(r2){recovery.push("Register restore failed: "+r2.message);}
            try{restoreLayer(oldNote,noteState);}catch(r3){recovery.push("note restore failed: "+r3.message);}
            try{removeStage(regStage);}catch(r4){recovery.push("Temporary Register cleanup failed: "+r4.message);}
            try{removeStage(noteStage);}catch(r5){recovery.push("Temporary note cleanup failed: "+r5.message);}
            if(problemUnits.length)try{focus=selectProblems();}catch(focusErr){recovery.push("Could not select related objects: "+focusErr.message);}
        }
        alert("Stopped.\n\n"+(err.message || String(err))+
              (focus ? "\n\n"+focus : "")+
              (recovery.length ? "\n\n"+recovery.join("\n") :
               (!committed ? "\n\nExisting artboards and output have been preserved." : "")));
    } finally {
        if(changedCS)try{app.coordinateSystem=oldCS;}catch(csErr){}
    }
})();
