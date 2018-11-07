const plugin = {};
(function(context) {

	/**
     * @param {number} wrap - is a value that specifies the content control type. It can have one of the following values: 1 (block) or 2 (inline)
     * @param {number} [lock = 3] -  is a value that defines if it is possible to delete and/or edit the content control or not. By default, content control can be edited and deleted
     * @param {number} [id = 0] -  is a identifier of the content control for search it in future. Fill free to set some value
     * @param {string} [tag = '0'] -  is a tag assigned to the content control.
     * @return Promise<ContentControl> object
     */
	function create_content_control(wrap, lock = 3, id = 0, tag = '0') {
        return new Promise(resolve => {
            window.Asc.plugin.executeMethod("AddContentControl", [wrap, {"Id": id, "Lock": lock, "Tag": tag}], x => {
                let result;
                if (x) {
                    result = new ContentControl(x)
                }
                resolve(result)
            })
        })
    }

    /**
     * @param {lock} [lock = 3] -  is a value that defines if it is possible to delete and/or edit the content control or not. By default, content control can be edited and deleted
     * @param {number} [id = 0] -  is a identifier of the content control for search it in future. Fill free to set some value
     * @param {string} [tag = '0'] -  is a tag assigned to the content control.
     * @return Promise<ContentControl> object
     */
	context.create_inline_content_control = function(lock = 3, id = 0, tag = '0') {
		if (typeof lock !== 'string' && typeof lock !== 'number') {
			id = lock.id || 0;
			tag = lock.tag || '0';
			lock = lock.lock || 3
        }
		return create_content_control(2, lock, id, tag)
    };


    /**
     * @return {Promise<ContentControl>} object
     */
	context.create_block_content_control = function() {
			if (typeof lock !== 'string' && typeof lock !== 'number') {
				id = lock.id || 0;
            	tag = lock.tag || '0';
            	lock = lock.lock || 3
        	}
		return this.create_content_control(1, lock, id, tag)
	};

	 /** is using for close plugin */
    context.close = function() {
        window.Asc.plugin.executeCommand("close", '');
    };

	/**
     * get content controls.
     * @return {Promise<Array>} with jsons. Example: [{"Tag":"","InternalId":"2_813"}]
      */
    context.get_all_content_controls = function() {
        return new Promise(resolve => {
            window.Asc.plugin.executeMethod("GetAllContentControls", [], content_controls => {
                resolve(content_controls.map(x => new ContentControl(x)
                ))
            })
        })
    };

	/**
     * get content controls by tag.
     * @return {Promise<Array>} with jsons. Example: [{"Tag":"","InternalId":"2_813"}]
     */
    context.get_content_controls_by_tag = function(tag) {
        return new Promise(resolve => {
            window.Asc.plugin.executeMethod("GetAllContentControls", [], content_controls => {
                let content_controls_with_tag = content_controls.filter(x => x['Tag'] === tag);
                resolve(content_controls_with_tag.map(x => new ContentControl(x)))
            })
        })
    };

	/**
     * get content controls by tag.
     * @return {Promise<Array>} with jsons. Example: [{"Tag":"","InternalId":"2_813"}]
     */
    context.get_content_control_by_id = function(id) {
        return new Promise(resolve => {
            window.Asc.plugin.executeMethod("GetAllContentControls", [], content_controls => {
                let content_controls_with_id = content_controls.filter(x => x['Id'] === id);
                if (content_controls_with_id.lenght > 1) {
                    console.warn('There are some content controls with same names. Only first of them will be return')
                }
                resolve(new ContentControl(content_controls_with_id[0]));
            })
        })
    };

	/**
     * get content control, if curwor is placed in it.
     * @return {ContentControl} object or undefined
     */
    context.get_current_content_control = function() {
        return new Promise(resolve => {
            window.Asc.plugin.executeMethod("GetCurrentContentControl", [], content_control => {
                if (content_control) {
                    resolve(new ContentControl({'InternalId': content_control}));
                } else {
                    resolve();
                }
            })
        })
    };

    /**
     * remove content control.
     * @param {Array} ids is array with strings id. Example: ['1', '2']
     * ids can be {String}, if you want to delete only one content control
     */
    context.remove_content_controls = function(ids) {
        if (typeof(ids) === "string") {
            ids = [ids]
        }
        return new Promise(resolve => {
            window.Asc.plugin.executeMethod("RemoveContentControls", [ids.map(id => {
                    return ({'InternalId': id})
                })], content_controls => {
                    resolve()
                }
            )
        })
    };

	/**
     * select content control
     * @param id
     */
    context.select_content_control = function(id) {
        return new Promise(resolve => {
            window.Asc.plugin.executeMethod("SelectContentControl", [id], callback => {
                    resolve()
                }
            )
        })
    };

	/**
     * select content control
     */
    context.get_selected_text = function() {
        return new Promise(resolve => {
            window.Asc.plugin.executeMethod("GetSelectedText", [], selected_text => {
                    resolve(selected_text)
                }
            )
        })
    };


	/**
     * method for adding data to content control.
     * @param {ContentControl} content_control object
     * @param {String} script is a data for runing
     * Please, use method  contentcontrol.run_script() for runnint script
     */
    function run_script_in_content_control(content_control, script) {
        return new Promise(resolve => {
            window.Asc.plugin.executeMethod("InsertAndReplaceContentControls", [
                    [
                        {
                            "Props":
                                {
                                    "Id": content_control.id,
                                    "Tag": content_control.tag,
                                    "Lock": content_control.lock
                                },
                            "Script": script
                        }
                    ]
                ], callback => {
                    resolve(callback);
                }
            )
        })
    }

	/**
     * method for adding data to content control.
     * @param {ContentControl} content_control object
     * @param {Number} lock is a number of lock type.
     */
    function change_content_control_lock(content_control, lock) {
        return new Promise(resolve => {
            window.Asc.plugin.executeMethod("InsertAndReplaceContentControls", [
                    [
                        {
                            "Props":
                                {
                                    "InternalId": content_control.internal_id,
                                    "Lock": lock
                                }

                        }
                    ]

                ], x => {
                    resolve(x)
                }
            )
        })
    }

	/**
     * method for adding text to content control.
     * @param {ContentControl} content_control
     * @param {String} text for adding to content control
     * Please, use contentcontrol.write_text() for adding text
     */
    function write_to_content_control(content_control, text) {
        return run_script_in_content_control(content_control, "var oDocument = Api.GetDocument(); \
                                                          var oParagraph = Api.CreateParagraph(); \
                                                          oParagraph.AddText('" + text.toString() + "'); \
                                                          oDocument.InsertContent([oParagraph]);")
    }

    let ContentControl = function(params) {
        let tag = params['Tag'];
        this.tag = () => tag;

        let id = params['Id'];
        this.id = () => id;

        let internal_id = params['InternalId'];
        this.internal_id = () => internal_id;

        let editing = (params['Lock'] === '2' || params['Lock'] === '3');
        this.editing = () => editing;

        let deleting = (params['Lock'] === '0' || params['Lock'] === '3');
        this.deleting = () => deleting;

        let lock = params['Lock'];


        /**
         * status is a ability to editing for content control
         * @param {Boolean} status
         * @return {Promise<ContentControl>}
         */
        this.edit_access = (status) => {
            let _lock;
            if (status) {
                _lock = deleting ? 3 : 2
            } else {
                _lock = deleting ? 0 : 1
            }

            return change_content_control_lock(this, _lock).then(x => {
                editing = status;
                lock = x[0]['Lock'];
            })
        };

        /**
         * status is a ability to delete content control
         * @param status
         */
        this.delete_access = (status) => {
            let _lock;
            if (status) {
                _lock = this._editing ? 3 : 0
            } else {
                _lock = this._editing ? 2 : 1
            }
            return change_content_control_lock(this, _lock).then(x => {
                deleting = status;
                lock = x[0]['Lock'];
            })
        };

        /** @param {Number | String} _lock is a code of access setting for content control. See https://api.onlyoffice.com/plugin/executemethod/insertandreplacecontentcontrols*/
        this.lock = (_lock) => {
            lock = _lock;
            editing = (_lock === '2' || _lock === '3');
            deleting = (_lock === '0' || _lock === '3');
        };

		/**
     	* get content control data
     	* @param {String} text is a data for adding in paragraph. Only one paragraph will be added to content control.
     	* All data in content control will be deleted before
     	*/
    	this.write_text = function(text) {
        	write_to_content_control(this, text);
    	};

    	/**
     	* run script for adding data to content control
     	* @param {String} script is a code for documentbuilder
     	*/
    	this.run_script = function(script) {
            context.run_script_in_content_control(this, script);
    	};

    	/** remove content control */
    	this.remove = function() {
            context.remove_content_controls(this.internal_id);
    	}
    }


})(plugin);