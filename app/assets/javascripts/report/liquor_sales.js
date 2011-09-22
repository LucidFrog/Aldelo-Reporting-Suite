function getLiquorSales(){
	start_date = $(".start_date").val();
	end_date = $(".end_date").val();
	liquor_type = $(".liquor_type").val();
	liquor_name = $(".liquor_name").val();
	$.ajax({url: '/report/liquor_sales',
		data:{start_date: start_date, end_date: end_date, liquor_type: liquor_type, liquor_name: liquor_name},
		success: function(liquor_sale_info){
			$("#liquor_sale_results").html(liquor_sale_info);
		}
	})
}