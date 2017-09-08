alertify.YoutubeDialog || alertify.dialog('YoutubeDialog', function () {
    var iframe;
    return {
        // dialog constructor function, this will be called when the user calls alertify.YoutubeDialog(videoId)
        main: function (videoId) {
            //set the videoId setting and return current instance for chaining.
            return this.set({
                'videoId': videoId
            });
        },
        // we only want to override two options (padding and overflow).
        setup: function () {
            return {
                options: {
                    //disable both padding and overflow control.
                    padding: !1,
                    overflow: !1,
                    title: 'Hướng dẫn',
                }
            };
        },
        // This will be called once the DOM is ready and will never be invoked again.
        // Here we create the iframe to embed the video.
        build: function () {
            // create the iframe element
            iframe = document.createElement('iframe');
            iframe.frameBorder = "no";
            iframe.width = "100%";
            iframe.height = "100%";
            // add it to the dialog
            this.elements.content.appendChild(iframe);

            //give the dialog initial height (half the screen height).
            this.elements.body.style.minHeight = screen.height * .5 + 'px';
        },
        // dialog custom settings
        settings: {
            videoId: undefined
        },
        // listen and respond to changes in dialog settings.
        settingUpdated: function (key, oldValue, newValue) {
            switch (key) {
                case 'videoId':
                    iframe.src = "https://www.youtube.com/embed/" + newValue + "?enablejsapi=1";
                    break;
            }
        },
        // listen to internal dialog events.
        hooks: {
            // triggered when the dialog is closed, this is seperate from user defined onclose
            onclose: function () {
                iframe.contentWindow.postMessage('{"event":"command","func":"pauseVideo","args":""}', '*');
            },
            // triggered when a dialog option gets update.
            // warning! this will not be triggered for settings updates.
            onupdate: function (option, oldValue, newValue) {
                switch (option) {
                    case 'resizable':
                        if (newValue) {
                            this.elements.content.removeAttribute('style');
                            iframe && iframe.removeAttribute('style');
                        } else {
                            this.elements.content.style.minHeight = 'inherit';
                            iframe && (iframe.style.minHeight = 'inherit');
                        }
                        break;
                }
            }
        }
    };
});
//show the dialog
alertify.YoutubeDialog('1m_9z7Kxhko').set({ frameless: false });

function notify() {
    console.log('notify');
    alertify.alert('Anh Chương có lời nhắn nhủ với bạn!',
        'Bạn chưa tải tệp lên mừ, làm sao mà tui xử lý đây!').set('label', 'Tôi đã hiểu rồi');
}

$('#phonenumberedit').off('click').on('click', notify);

$('#UploadedFile').change(function () {
    if ($(this).val().length > 0) {
        $('#form').submit();
    }
});

$('#form').on('submit', function (e) {
    e.preventDefault();

    var data = new FormData(this);

    $.ajax({
        type: 'post',
        url: '/filterdata/UploadPhoneNumber',
        data: data,
        success: function (res) {
            console.log(res.status);
            if (res.status == false && res.large == false) {
                alertify.alert('Anh Chương có lời nhắn nhủ với bạn!',
                    'Bạn chưa chọn tệp hoặc là tệp không đúng định dạng, bạn vui lòng kiểm tra lại trước khi tải lên.').set('label', 'Tôi đã hiểu rồi');

                $('#phonenumberedit').off('click').on('click', notify);
            }
            if (res.status == false && res.large == true) {
                alertify.alert('Anh Chương có lời nhắn nhủ với bạn!', res.data + '<br><a href="/filterdata/DownloadPhoneNumber"> Vui lòng click vào đây để lọc dữ liệu</a>').set('label', 'Tôi đã hiểu rồi');
                //window.location.href = '/filterdata/DownloadPhoneNumber';
            }
            else {
                var data = res.data;
                var rendered = '';
                var template = $('#template').html();

                var stt = 1;
                //for (var d in data) {
                //    rendered += Mustache.render(template, {
                //        Id: stt++,
                //        HoTen: d.HoTen,
                //        DienThoai: d.DienThoai,
                //        DiaChi: d.DiaChi,
                //    });
                //}
                $.each(data, function (i, item) {
                    rendered += Mustache.render(template, {
                        Id: i + 1,
                        HoTen: item.HoTen,
                        DienThoai: item.DienThoai,
                        DiaChi: item.DiaChi,
                    });
                });
                $('#target').html(rendered);

                $('#phonenumberedit').off('click').on('click', function () {
                    console.log('download');
                    window.location.href = '/filterdata/DownloadPhoneNumber';
                });
            }
        },
        cache: false,
        contentType: false,
        processData: false
    })
})