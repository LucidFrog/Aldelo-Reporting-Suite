function getOvertime(){
	overtime_date = $(".datepicker_input").val();
	job_title_id = $(".job_title_id").val();
	$.ajax({url: '/report/overtime',
		data:{overtime_date: overtime_date, job_title_id: job_title_id},
		success: function(overtime_info){
			$("#overtime_results").html(overtime_info);
		}
	})
}