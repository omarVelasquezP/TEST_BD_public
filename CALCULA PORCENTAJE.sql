declare

@mto_sum money,
@num_sum int,
@vip_porcentaje int



set @mto_sum = 1000
set @num_sum = 10
set @vip_porcentaje = 50
select @mto_sum as mto_sum,@num_sum as num_sum,@vip_porcentaje as vip_porcentaje

set @mto_sum = @mto_sum + (@mto_sum * (convert(money,@vip_porcentaje)/100))
set @num_sum = @num_sum + (@num_sum * (convert(money,@vip_porcentaje)/100))

select @mto_sum as mto_sum,@num_sum as num_sum
