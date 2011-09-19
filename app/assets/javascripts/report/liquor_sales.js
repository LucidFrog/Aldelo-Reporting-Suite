function getLiquorSales(){
	start_date = $(".datepicker_input").val();
	$.ajax({url: '/report/liquor_sales',
		data:{start_date: start_date},
		success: function(liquor_sale_info){
			$("#liquor_sale_results").html(liquor_sale_info);
		}
	})
}