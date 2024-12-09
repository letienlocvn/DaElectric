package model;

import lombok.*;

import java.util.Date;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class Customer {
    // Mã số khách hàng (unique identifier)
    private int customerId;

    // Họ và tên khách hàng
    private String fullName;

    // Chỉ số điện cũ
    private int oldIndex;

    // Chỉ số điện mới
    private int newIndex;

    // Số điện tiêu thụ trong tháng
    private int consumedElectricity;

    // Đơn giá điện
    private double unitPrice;

    // Thành tiền điện
    private double electricityCost;

    // Công ghi điện (phí dịch vụ)
    private double meterFee;

    // Tổng thanh toán
    private double totalPayment;

    // Địa chỉ khách hàng
    private String address;

    // Số điện thoại khách hàng
    private String phoneNumber;

    // Ngày lập hóa đơn
    private Date billDate;

    @Override
    public String toString() {
        return "Customer{" +
                "customerId=" + customerId +
                ", fullName='" + fullName + '\'' +
                ", oldIndex=" + oldIndex +
                ", newIndex=" + newIndex +
                ", consumedElectricity=" + consumedElectricity +
                ", unitPrice=" + unitPrice +
                ", electricityCost=" + electricityCost +
                ", meterFee=" + meterFee +
                ", totalPayment=" + totalPayment +
                ", address='" + address + '\'' +
                ", phoneNumber='" + phoneNumber + '\'' +
                ", billDate=" + billDate +
                '}';
    }
}
