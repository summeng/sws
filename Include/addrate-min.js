KISSY.add("rate/addrate",function(a){function b(b,c,d){var e=this;e.el=a.one(b),e.el_share=e.el.one(".rate-share"),e.el_annoy=e.el.one(".rate-annoy"),e.id=e.el.attr("data-id"),e.form=c,e.el_msg_box=e.el.one(".rate-msg-box"),e.el_msg_input=e.el.one(".rate-msg"),e.el_ctl_star=e.el.one(".rate-control-stars"),e.el_tips=e.el.one("div.gift-tip"),e.el_stars=e.el.one("span.rate-stars"),e.ratecontrol=e.el.one(".rate-controll"),e.config=d,e.defaultShow=e.el.hasClass("st-show-msg-box"),e.defaultShow?e.showComment():e.hideComment(),new g.PlaceHolder(e.el_msg_input[0],{elLabel:e.el.one(".ph-label")[0]});var f=new g.WordCounter(e.el_msg_input);f.on("countupdate",function(a){var b=e.el_msg_box.one(".r-t-counter"),c=500-a.count;b.html(Math.max(c,0))}),e.el.on("click",function(b){var c=a.one(b.target);c.hasClass("show-msg-box")?(b.halt(),e.showComment(!0)):c.hasClass("close-msg-box")&&(b.halt(),e.hideComment())}),e.on("dsrchange",function(){e.showComment()}),e.initRateControl(),e.el_annoy&&e.el_annoy.on("click",function(){e.fire("annoychange",{checked:e.el_annoy[0].checked})}),d.isB2C&&(e.isB2C=!0,e.initSimpleStar(e.el_stars,d)),e.on("annoychange",function(){var a=e.el_share[0];e.el_annoy[0].checked?(a.disabled=!0,a.checked=!1):a.disabled=!1}),e.on("score_change",function(){e.showComment()})}function c(c){function f(){var b=a.one("#starstip");b&&b.hide(),a.each(n,function(a){a.detach("change",f)})}var j,k=[],l=a.all(".rate-box"),m=a.one("#rateListForm"),n=[];l.each(function(d,e){k.push(new b(d,m,c)),e===l.length-1&&a.one(d).addClass("rate-box-last")}),a.Event.on(k,"ratechange",function(){a.one("#rate-good-all")[0].checked=!1});var o=a.one("#annoy-all"),p=!1;d.on(k,"annoychange",function(a){a.checked||(o[0].checked=!1)}),o&&o.on("click",function(){var b=a.one(this)[0].checked;a.each(k,function(a){a.setAnnoy(b)}),p=!0});var q=a.one("#rate-good-all");q&&q.on("click",function(){var b=a.one(this)[0].checked;a.each(k,function(a){a.setGoodRate(b)})});var r=a.one("ul.stars");r&&r.all(".rate-stars").length>0&&r.all(".rate-stars").each(function(b){var d=a.one(b).attr("data-type"),e=d?c.star_tooltip[d]:[],f=new g.SimpleStar(b,{tooltips:e,levelText:i});n.push(f)}),a.each(n,function(a){a.on("change",f)}),m.on("submit",function(b){var d,f=!1,g=0,i=0,l=!1,o=!1,p=0,q=0;if(b.halt(),a.each(k,function(a){a.hideError();var b=a.checkForm();f=f||b}),!f){if(a.each(k,function(a){a.rateradios&&(p++,null!==a.getRateScore()&&q++),a.dsr&&(g++,-1!==a.dsr.index()&&i++)}),a.each(n,function(a){g++,-1!==a.index()&&i++}),o=p>0&&0==q,l=g>0&&0==i,o&&l||0==p&&l||0==g&&o)return d=m.one("div.submitbox").one("div.msg"),d||(d=a.one(e.create('<div class="msg">')),d.insertBefore(m.one("div.submitbox").one("button"))),void a.later(function(){d.html('<p class="error">\u4f60\u8fd8\u6ca1\u6709\u6dfb\u52a0\u4efb\u4f55\u8bc4\u4ef7,\u4e0d\u8981\u5077\u61d2\u54e6\uff0c\u4eb2!</p>').show()},50);if(!(c.isB2C&&g>0&&g>i)||confirm("\u4f60\u8fd8\u6709\u672a\u5b8c\u6210\u7684\u8bc4\u5206\uff0c\u786e\u5b9a\u4ee5\u540e\u90fd\u4e0d\u8bc4\u5206\u4e86\uff1f")){h.LoginCheck||m[0].submit();var r=m.one("button.submit"),s=r.html();r.prop("disabled",!0).html("\u63d0\u4ea4\u4e2d..."),j||(j=new h.LoginCheck({ajaxURL:m.attr("data-checkLoginJson"),callbackURL:m.attr("data-checkLoginCallback"),success:function(b){a.isObject(b)&&a.isArray(b.update)&&a.each(b.update,function(a){var b=m[0].elements[a.name];b&&(b.value=a.value)}),r.html(s),m[0].submit()},cancel:function(){r.html(s).prop("disabled",!1)}})),j.start()}}})}var d=a.Event,e=a.DOM,f=a.namespace("rate"),g=a.namespace("ui"),h=a.namespace("util"),i=["\u5f88\u4e0d\u6ee1\u610f","\u4e0d\u6ee1\u610f","\u4e00\u822c","\u6ee1\u610f","\u5f88\u6ee1\u610f"],j="st-show-msg-box";a.augment(b,a.EventTarget,{initRateControl:function(){var b=this;b.ratecontrol&&(b.rateradios=b.ratecontrol.all("input"),b.rateradios&&(b.rateradios.on("click",function(c){a.one(c.target);b.fire("ratechange"),b.initRateScore()}),b.initRateScore()))},initRateScore:function(){var a,b,c=this,d="",e=c.el.one(".rate-score"),f=c.config.placeholder,g=f.BAD,h=f.NORMAL,i=c.el.attr("data-prompt")||f.GOOD;if(a=c.getRateScore(),"number"==typeof a){switch(a){case-1:b=g,d="bad";break;case 1:b=i,d="good";break;case 0:b=h,d="normal"}0>=a?c.el.addClass("rate-level-normal"):c.el.removeClass("rate-level-normal"),a=a>0?"+"+a:""+a,e.show().html(a+"\u5206").attr("class","rate-score "+d),"0"===a&&e.show().html("\u4e0d\u52a0\u5206"),c.el.one(".ph-label").html(b),c.fire("score_change")}else e.hide().html("&nbsp;")},initSimpleStar:function(a,b){var c=this;c.dsr=new g.SimpleStar(a,{tooltips:b.star_tooltip.match,levelText:i,el_tip:c.el.one(".dsr-score"),tipstmpl:"<em>{score}</em> \u5206"}),c.dsr.on("change",function(){c.fire("dsrchange")})},checkForm:function(){var b=this,c=b.el_msg_input,d=a.trim(c.val()),e=b.getRateScore();return error="",b.isB2C?b.dsr&&-1==b.dsr.index()&&d&&(error="\u6709\u8bc4\u8bba\u5185\u5bb9\u7684\u5b9d\u8d1d\u9700\u8981\u6253\u5206",this.showComment()):(d&&null===e&&(error="\u53d1\u8868\u8bc4\u8bba\u7684\u5b9d\u8d1d\u5fc5\u987b\u9009\u62e9\u597d\u3001\u4e2d\u3001\u5dee\u8bc4",this.showComment(!0)),null!==e&&0>=e&&!d&&(error="\u4e2d\u5dee\u8bc4\u4e00\u5b9a\u8981\u8bf4\u660e\u539f\u56e0\u54e6",c[0].focus())),error?(b.showError(error),document.body.scrollTop=b.el.offset().top,error):void 0},showError:function(b){var c=this,d=c.el_msg_box.one(".msg");d||(d=a.one(e.create('<div class="msg">')),d.insertBefore(c.el_msg_box.one(".textinput"))),d.html('<p class="stop">'+b+"</p>").show()},hideError:function(){var a=this,b=a.el_msg_box.all(".msg");b&&b.length>0&&b.hide()},showComment:function(b){var c=this;c.el.one("a.show-msg-box").hide(),c.el_msg_box.slideDown(.2,function(){c.el.addClass(j),c.el_tips&&c.el_tips.show(),b&&a.later(function(){c.el_msg_input[0].focus()},100)})},hideComment:function(){var b=this;return ratescore=b.getRateScore(),msg=a.trim(b.el_msg_input.val()),label=msg?"\u4fee\u6539\u8bc4\u8bba":"\u6dfb\u52a0\u8bc4\u8bba",null!==ratescore&&1>ratescore&&""===msg?void b.el_msg_input[0].focus():(b.el.removeClass(j),b.el_tips&&b.el_tips.hide(),void b.el_msg_box.slideUp(.2,function(){b.el.one("a.show-msg-box").show().html(label)}))},getRateScore:function(){var b,c=this,d=c.rateradios;return d&&d.each(function(a){a[0].checked&&(b=a.val())}),c.el_stars&&c.el_stars.all("input").each(function(a){a[0].checked&&(b=a.val())}),""===a.trim(b)?null:parseInt(b,10)},setAnnoy:function(a){if(this.el_annoy){var b=this,c=b.el_annoy[0];c.disabled||c.checked!=a&&(c.checked=a,b.fire("annoychange",{checked:a}))}},setGoodRate:function(a){this.el.one(".good-rate")[0].checked=a,this.initRateScore()}}),f.rateInit=c});