function getTotalHours(){
	total_hours_date = $(".datepicker_input").val();
	job_title_id = $(".job_title_id").val();
	$.ajax({url: '/report/total_hours',
		data:{total_hours_date: total_hours_date, job_title_id: job_title_id},
		success: function(total_hours_info){
			$("#total_hours_results").html(total_hours_info);
		}
	})
}