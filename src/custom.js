/* Multi level dropdown support */

$(function () {
  $("ul.dropdown-menu [data-toggle='dropdown']").on("click", function (event) {
    event.preventDefault();
    event.stopPropagation();

    // Reset other active dropdowns
    $("ul.dropdown-menu .show").removeClass("show");

    $(this).siblings().toggleClass("show");

    if (!$(this).next().hasClass("show")) {
      $(this)
        .parents(".dropdown-menu")
        .first()
        .find(".show")
        .removeClass("show");
    }

    $(this)
      .parents("li.nav-item.dropdown.show")
      .on("hidden.bs.dropdown", function (e) {
        $(".dropdown-submenu .show").removeClass("show");
      });
  });
});

/* Tab table switcher */

$(function () {
  $(".table-tabs-selector button").on("click", function (event) {
    $(".table-tabs-selector button").removeClass("active");
    $(this).addClass("active");
    var type = $(this).data("type");
    var target = $(this).data("target");
    $("#" + target).removeClass();
    $("#" + target).addClass(type);
  });
});


/* Loading functionality */

function showLoader(message) {
  if (message) {
    $(".loading-message").text(message);    
  }
  $(".loader").addClass("show");  
}

function hideLoader() {
  $(".loader").removeClass("show");
}
