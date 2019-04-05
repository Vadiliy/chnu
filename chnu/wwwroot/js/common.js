// SAME WIDTH FUNCTION
function sameWidth(selector) {
    var max_width = 0;
    $(selector).each(function () {
        if (max_width < $(this).width()) {
            max_width = $(this).width();
        }
    });
    $(selector).each(function () {
        $(this).width(max_width);
    });
};

$(document).ready(function () {
    // datepicker
    $( '#date').datepicker({
        changeMonth: true,
        dateFormat: "dd.mm.yy"
    });

    sameWidth('.byDiscipline .retake-item .name');
    sameWidth('.byDiscipline .retake-item .discipline-teacher');
    sameWidth('.byGroup .retake-item .student');
    sameWidth('.byGroup .retake-item .discipline-info');
    sameWidth('.byName .retake-item .discipline-info');
    sameWidth('.admin .retake-item .student .name');
    sameWidth('.admin .retake-item .student .group');

    $('*[data-toggle]').on('click', function () {
        var toggle = $(this).data('toggle');
        $('.' + toggle).toggleClass('show');
    });

    //
    $('*[data-confirm]').on('click', function () {
        var form = $(this).data('confirm');
        $('form#' + form).submit();
    });

    // OPEN MODALS
    $('*[data-modal]').on('click', function () {
        
        var stId = $(this).parent().parent().find(".student").find(".student-id").text();
        $('<input>').attr({
            type: 'hidden',
            id: 'studentId',
            name: 'studentId',
            value: stId
        }).appendTo('div .input-group');

        var discId = $(this).parent().parent().find(".discipline-info").find(".discipline-id").text();
        $('<input>').attr({
            type: 'hidden',
            id: 'disciplineId',
            name: 'disciplineId',
            value: discId
        }).appendTo('div .input-group');


        $('body').addClass('modal-open');
        var modal = $(this).data('modal');
        $('.' + modal).fadeIn().addClass('flex');

        
    });

    // CLOSE MODALS
    $('.modal-container .close-btn, .modal-container .overlay').on('click', function () {
        $('body').removeClass('modal-open');
        $('.modal-container').fadeOut(function () {
            $('.modal-container').removeClass('flex');
        });
        console.log('click');
    });

    //  FAKE SELECT
    $('.fakeSelect').on('click', function () {
        $(this).children('.selectOption').slideToggle();
    });
    $('.selectOption .option p').on('click', function () {
        var text = $(this).html();
        var selectOption = $(this).data('select');
        var selectTo = $(this).closest('.fakeSelect').data('select-to');
        $(this).closest('.selectOption ').siblings('.selected').children('p').html(text);
        $("select#" + selectTo + " option").attr('value', selectOption).attr('selected', true).html(text);
    });
});
