/**
 * This jQuery plugin displays pagination links inside the selected elements.
 *
 * @author Gabriel Birke (birke *at* d-scribe *dot* de)
 * @version 1.1
 * @param {int} maxentries Number of entries to paginate
 * @param {Object} opts Several options (see README for documentation)
 * @return {Object} jQuery Object
 */
jQuery.fn.pagination = function (maxentries, opts) {
    opts = jQuery.extend({
        items_per_page: 10,
        num_display_entries: 10,
        current_page: 0,
        num_edge_entries: 0,
        link_to: "javascript:void(0);",
        first_text: "First",//Radys Add this link  more info  http://radys.cn
        last_text: "Last", //Radys Add this link
        prev_text: "Prev",
        next_text: "Next",
        goto: "Go",
        goto_name: "txtGoto",
        ellipse_text: "...",
        prev_show_always: true,
        next_show_always: true,
        callback: function () { return false; }
    }, opts || {});

    return this.each(function () {
		/**
		 * Calculate the maximum number of pages
		 */
        function numPages() {
            return Math.ceil(maxentries / opts.items_per_page);
        }

		/**
		 * Calculate start and end point of pagination links depending on 
		 * current_page and num_display_entries.
		 * @return {Array}
		 */
        function getInterval() {
            var ne_half = Math.ceil(opts.num_display_entries / 2);
            var np = numPages();
            var upper_limit = np - opts.num_display_entries;
            var start = current_page > ne_half ? Math.max(Math.min(current_page + 1 - ne_half, upper_limit), 0) : 0;
            var end = current_page > ne_half ? Math.min(current_page + ne_half, np) : Math.min(opts.num_display_entries, np);
            return [start, end];
        }

		/**
		 * This is the event handling function for the pagination links. 
		 * @param {int} page_id The new page number
		 */
        function pageSelected(page_id, evt, flag) {
            if (flag == 1) {
                $.each($("input:text"), function (i, val) {
                    if (val.id == opts.goto_name) {
                        if ($("#" + opts.goto_name).val() != "") {
                            page_id = parseInt($("#" + opts.goto_name).val()) - 1;
                            if (page_id < 0) { page_id = 0; }
                        }
                        else {
                            page_id = current_page;
                        }
                    }
                });
            }

            current_page = page_id;
            drawLinks();
            var continuePropagation = opts.callback(page_id, panel);
            if (!continuePropagation) {
                if (evt.stopPropagation) {
                    evt.stopPropagation();
                }
                else {
                    evt.cancelBubble = true;
                }
            }
            return continuePropagation;
        }

		/**
		 * This function inserts the pagination links into the container element
		 */
        function drawLinks() {
            var nGotoPage = 1;
            $.each($("input:text"), function (i, val) {
                if (val.id == opts.goto_name) {
                    if ($("#" + opts.goto_name).val() != "") {
                        nGotoPage = parseInt($("#" + opts.goto_name).val());
                    }
                    else {
                        nGotoPage = current_page;
                    }
                }
            });

            if (nGotoPage > np) { nGotoPage = np; }
            if (nGotoPage < 1) { nGotoPage = 1; }

            panel.empty();
            var interval = getInterval();
            var np = numPages();

            // Helper function for generating a single link (or a span tag if it'S the current page)
            var appendItem = function (page_id, appendopts) {
                page_id = page_id < 0 ? 0 : (page_id < np ? page_id : np - 1); // Normalize page id to sane value
                appendopts = jQuery.extend({ text: page_id + 1, classes: "" }, appendopts || {});
                if (page_id == current_page && appendopts.text != opts.goto) {
                    var lnk = $("<a class='page-link'>" + (appendopts.text) + "</a>").attr('href', opts.link_to);
                }
                else {
                    if (appendopts.text == opts.goto) {
                        var lnk = $("<a class='page-link'>" + (appendopts.text) + "</a>")
                            .bind("click", getClickHandler(page_id, 1))
                            .attr('href', opts.link_to.replace(/__id__/, page_id));
                    }
                    else {
                        var lnk = $("<a class='page-link'>" + (appendopts.text) + "</a>")
                            .bind("click", getClickHandler(page_id, 0))
                            .attr('href', opts.link_to.replace(/__id__/, page_id));
                    }
                }

                lnk = $("<li class='page-item'></li>").append(lnk);
                if (appendopts.classes) lnk.addClass(appendopts.classes);
                if (page_id == current_page && !appendopts.classes) lnk.addClass("active");
                if (current_page == 0 && appendopts.classes.indexOf('previous') >= 0) lnk.addClass("disabled");
                if (current_page == (np - 1) && appendopts.classes.indexOf('next') >= 0) lnk.addClass("disabled");

                panel.append(lnk);
            }

            // This helper function returns a handler function that calls pageSelected with the right page_id
            var getClickHandler = function (page_id, flag) {
                return function (evt) { return pageSelected(page_id, evt, flag); }
            }

            if (np > 1) {
                // Radys Add
                // Generate "First"-Link
                if (opts.first_text && (current_page > 0 || opts.prev_show_always)) {
                    appendItem(0, { text: opts.first_text, classes: "previous" });
                }

                // Generate "Previous"-Link
                if (opts.prev_text && (current_page > 0 || opts.prev_show_always)) {
                    appendItem(current_page - 1, { text: opts.prev_text, classes: "previous" });
                }

                // Generate starting points
                if (interval[0] > 0 && opts.num_edge_entries > 0) {
                    var end = Math.min(opts.num_edge_entries, interval[0]);
                    for (var i = 0; i < end; i++) {
                        appendItem(i);
                    }
                    if (opts.num_edge_entries < interval[0] && opts.ellipse_text) {
                        $("<li class='page-item disabled'><a class='page-link'>" + opts.ellipse_text + "</a></li>").appendTo(panel);
                    }
                }

                // Generate interval links
                for (var i = interval[0]; i < interval[1]; i++) {
                    appendItem(i);
                }

                // Generate ending points
                if (interval[1] < np && opts.num_edge_entries > 0) {
                    if (np - opts.num_edge_entries > interval[1] && opts.ellipse_text) {
                        jQuery("<li class='page-item disabled'><a class='page-link'>" + opts.ellipse_text + "</a></li>").appendTo(panel);
                    }
                    var begin = Math.max(np - opts.num_edge_entries, interval[1]);
                    for (var i = begin; i < np; i++) {
                        appendItem(i);
                    }
                }

                // Generate "Next"-Link
                if (opts.next_text && (current_page < np - 1 || opts.next_show_always)) {
                    appendItem(current_page + 1, { text: opts.next_text, classes: "next" });
                }

                // Radys Add
                // Generate "Last"-Link
                if (opts.last_text && (current_page < np - 1 || opts.next_show_always)) {
                    appendItem(np - 1, { text: opts.last_text, classes: "next" });
                }

                //text box goto
                //var sGoTo = "<a style='height:31px;'><input type='text' id='" + opts.goto_name + "' style='width:20px;height:20px;border:0px;background-color: #fff;' value='" + nGotoPage + "' onkeydown='if(((event.keyCode>47)&&(event.keyCode<58))||((event.keyCode>95)&&(event.keyCode<106))||(event.keyCode==8)||(event.keyCode==46)){return true;}else{return false;}'/></a>";
                //panel.append(sGoTo);
                //appendItem(nGotoPage - 1, { text: opts.goto, classes: "" });
            }
        }

        // Extract current_page from options
        var current_page = opts.current_page;
        // Create a sane value for maxentries and items_per_page
        maxentries = (!maxentries || maxentries < 0) ? 1 : maxentries;
        opts.items_per_page = (!opts.items_per_page || opts.items_per_page < 0) ? 1 : opts.items_per_page;
        // Store DOM element for easy access from all inner functions
        var panel = jQuery(this);
        // Attach control functions to the DOM element 
        this.selectPage = function (page_id) { pageSelected(page_id); }
        this.prevPage = function () {
            if (current_page > 0) {
                pageSelected(current_page - 1);
                return true;
            }
            else {
                return false;
            }
        }
        this.nextPage = function () {
            if (current_page < numPages() - 1) {
                pageSelected(current_page + 1);
                return true;
            }
            else {
                return false;
            }
        }

        var nFlag = 0;
        // When all initialisation is done, draw the links
        drawLinks();
    });
}