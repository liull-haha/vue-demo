var shop = new Vue({
    el: "#shop",
    data: {
        totalMoney: 0,
        goodList: []
    },
    filters: {

    },
    mounted: function () {
        this.selectAll();
    },
    methods: {
        selectAll: function () {
            var _this = this;
            this.$http.get("../data/goodsinfo").then(function (res) {
                _this.totalMoney = res.body.totalMoney;
                _this.goodList = res.body.goodList;
            })
        }
    }
});