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
alertify.YoutubeDialog('XJ3GS0bRtP0').set({ frameless: false });



//alertify.alert('Anh Chương có lời nhắn nhủ với bạn!',
//    'Bạn chưa tải tệp lên mừ, làm sao mà tui xử lý đây!').set('label', 'Tôi đã hiểu rồi');

function notify() {
    console.log('notify');
    alertify.alert('Anh Chương có lời nhắn nhủ với bạn!',
        'Bạn chưa tải tệp lên mừ, làm sao mà tui xử lý đây!').set('label', 'Tôi đã hiểu rồi');
}

$('#download').off('click').on('click', notify);
$('#download1').off('click').on('click', notify);

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
        url: '/home/upload',
        data: data,
        success: function (res) {
            console.log(res.status);
            if (res.status == false) {
                alertify.alert('Anh Chương có lời nhắn nhủ với bạn!',
                    'Bạn chưa chọn tệp hoặc là tệp không đúng định dạng, bạn vui lòng kiểm tra lại trước khi tải lên.').set('label', 'Tôi đã hiểu rồi');

                $('#download').off('click').on('click', notify);
                $('#download1').off('click').on('click', notify);
            }
            else {
                var data = res.data;
                var rendered = '';
                var template = $('#template').html();
                $.each(data, function (i, item) {
                    rendered += Mustache.render(template, {
                        Id: i + 1,
                        HoTen: item.HoTen,
                        DienThoai: item.DienThoai,

                        ThoiGian1: item.ThoiGian1,
                        Ca1: item.Ca1,
                        ThoiGian2: item.ThoiGian2,
                        Ca2: item.Ca2,
                        ThoiGian3: item.ThoiGian3,
                        Ca3: item.Ca3,
                        KickSale: item.KickSale,
                        Ngay: item.Ngay
                    });
                });
                $('#target').html(rendered);

                $('#download').off('click').on('click', function () {
                    console.log('download');
                    window.location.href = '/home/download';
                });
                $('#download1').off('click').on('click', function () {
                    console.log('download1');
                    window.location.href = '/home/download1';
                });
            }
        },
        cache: false,
        contentType: false,
        processData: false
    })
})