/**
 * GOWordAgent WPS 加载项入口
 */

// 初始化
window.onload = function() {
    console.log('GOWordAgent WPS Addon loading...');
    
    UIController.init();
    ProofreadWorkflow.init();
    
    console.log('GOWordAgent WPS Addon loaded');
};

// WPS 加载项生命周期回调
var wpsAddon = {
    onOpen: function() {
        console.log('Addon opened');
    },
    
    onClose: function() {
        console.log('Addon closed');
    }
};
