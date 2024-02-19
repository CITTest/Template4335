namespace Template4335.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class SHIT : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.Orders", "Status", c => c.String());
            AddColumn("dbo.Orders", "ClosingDate", c => c.DateTime(nullable: false));
            AddColumn("dbo.Orders", "RentTime", c => c.DateTime(nullable: false));
        }
        
        public override void Down()
        {
            DropColumn("dbo.Orders", "RentTime");
            DropColumn("dbo.Orders", "ClosingDate");
            DropColumn("dbo.Orders", "Status");
        }
    }
}
