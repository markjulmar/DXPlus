namespace DXPlus
{
    /// <summary>
    /// Extension methods for the Picture/Image type
    /// </summary>
    public static class PictureExtensions
    {
        /// <summary>
        /// Fluent method to set rotation
        /// </summary>
        /// <param name="picture">Picture</param>
        /// <param name="value">Rotation</param>
        /// <returns>Picture object</returns>
        public static Picture SetRotation(this Picture picture, int value)
        {
            picture.Rotation = value;
            return picture;
        }

        /// <summary>
        /// Flip the picture horizontally
        /// </summary>
        /// <param name="picture">Picture</param>
        /// <param name="value">True/False</param>
        /// <returns>Picture object</returns>
        public static Picture FlipHorizontal(this Picture picture, bool value)
        {
            picture.FlipHorizontal = value;
            return picture;
        }

        /// <summary>
        /// Flip the picture vertically
        /// </summary>
        /// <param name="picture">Picture</param>
        /// <param name="value">True/False</param>
        /// <returns>Picture object</returns>
        public static Picture FlipVertical(this Picture picture, bool value)
        {
            picture.FlipVertical = value;
            return picture;
        }

        /// <summary>
        /// Method to set the width of a picture
        /// </summary>
        /// <param name="picture">Picture</param>
        /// <param name="value">New width</param>
        /// <returns>Picture object</returns>
        public static Picture SetWidth(this Picture picture, int value)
        {
            picture.Width = value;
            return picture;
        }

        /// <summary>
        /// Method to set the height of a picture
        /// </summary>
        /// <param name="picture">Picture</param>
        /// <param name="value">New width</param>
        /// <returns>Picture object</returns>
        public static Picture SetHeight(this Picture picture, int value)
        {
            picture.Height = value;
            return picture;
        }

        /// <summary>
        /// Method to set the name of a picture
        /// </summary>
        /// <param name="picture">Picture</param>
        /// <param name="value">New width</param>
        /// <returns>Picture object</returns>
        public static Picture SetName(this Picture picture, string value)
        {
            picture.Name = value;
            return picture;
        }

        /// <summary>
        /// Method to set the description of a picture
        /// </summary>
        /// <param name="picture">Picture</param>
        /// <param name="value">New width</param>
        /// <returns>Picture object</returns>
        public static Picture SetDescription(this Picture picture, string value)
        {
            picture.Description = value;
            return picture;
        }

        /// <summary>
        /// Method to set the decorative status of the picture
        /// </summary>
        /// <param name="picture">Picture</param>
        /// <param name="value">New value</param>
        /// <returns>Picture object</returns>
        public static Picture SetDecorative(this Picture picture, bool value)
        {
            picture.IsDecorative = value;
            return picture;
        }
    }
}
